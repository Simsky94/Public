using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DVAModelsReflection;
using System.Data;
using DVAModelsReflectionFINA.Models.FINA;
//using DVAModelsReflectionFINA1.Models.FINA;
using DVAModelsReflection.Models.CONT;
using DVAModelsReflection.Models.PSI;
using DVAModelsReflectionFINA1.Models.FINA;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    class ProcesoDeLLenadoTablaRM
    {
        DB2Database _db = null;
        int IdAgencia = 0;
        int anio = 0;
        int mes = 0;
        List<ConceptosContables> liConceptos = new List<ConceptosContables>();
        string query = "";

        public ProcesoDeLLenadoTablaRM(DB2Database _db, int aidAgencia, int mes, int anio)
        {
            this._db = _db;
            this.IdAgencia = aidAgencia;
            this.mes = mes;
            this.anio = anio;
            liConceptos = ConceptosContables.Listar(_db);
            Console.WriteLine("[INICIA RM]Inicia el proceso de llenado para la agencia ID: " + aidAgencia);
        }

        public void SumaFacturas(DB2Database _db, ref decimal[] facTaller, ref decimal[] facHYP, int aMes, int aAnio, int aIdAgencia)
        {
            string docAnt = "";
            int factor = 0;

            DataSet ds = ProcesoResultadoMensual.DevolverDatosOrdenado(_db, aMes, aAnio, aIdAgencia);

            ds = DevolverDatosOrdenado(ds);



            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                if (docAnt != ds.Tables[0].Rows[i]["FISENDCCVN"].ToString().Trim())
                {
                    docAnt = ds.Tables[0].Rows[i]["FISENDCCVN"].ToString().Trim();
                    if (ds.Tables[0].Rows[i]["FSSETPDCVN"].ToString() == "F")
                        factor = 1;
                    else
                    {
                        if (Convert.ToInt32(ds.Tables[0].Rows[i]["FICAIDTINC"]) == 3) //CA_CA_ID_TIPO_NOTA_CREDITO_CANCELACION
                        {
                            factor = -1;
                        }
                        else
                        {
                            factor = 0;
                        }
                    }
                    if ((ds.Tables[0].Rows[i]["FISEIDTIPP"].ToString() == "1" &&
                        ds.Tables[0].Rows[i]["TIPOORDEN"].ToString() != "22" &&
                        ds.Tables[0].Rows[i]["TIPOORDEN"].ToString() != "26") ||
                        ds.Tables[0].Rows[i]["FISEIDTIPP"].ToString() == "2" ||
                        ds.Tables[0].Rows[i]["FISEIDTIPP"].ToString() == "3" ||
                        ds.Tables[0].Rows[i]["FISEIDTIPP"].ToString() == "4")
                        facTaller[Convert.ToInt16(ds.Tables[0].Rows[i]["FISEIDTIPP"]) - 1] += factor;
                    else
                    {
                        facHYP[Convert.ToInt16(ds.Tables[0].Rows[i]["FISEIDTIPP"]) - 1] += factor;
                    }
                }
            }
        }

        public DataSet DevolverDatosOrdenado(DataSet dsOrigen)
        {
            DataSet dsTemp = new DataSet();

            DataView dv = new DataView(dsOrigen.Tables[0],
                                       "FISEIDCNVN > 0",
                                       "FSSETPDCVN ASC, FISENDCCVN ASC", DataViewRowState.CurrentRows);
            dsTemp.Tables.Add(dv.ToTable());
            return dsTemp;
        }

        public void InsertaRegistros()
        {
            //Todo aqui
            try
            {

                DateTime dateIni = DateTime.Now;
                int total = 0;

                decimal[] facTaller = new decimal[5];
                decimal[] facHYP = new decimal[5];

                for (int i = 0; i < 5; i++)
                    facTaller[i] = facHYP[i] = 0;

                bool diaPrimero = false; // regresar a false al liberar
                if (dateIni.Day == 1)
                {
                    diaPrimero = true;
                }

                //diaPrimero = false;

                List<ProcesoResultadoMensualExtralibros> RMDataCalculos9Extralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos9Extralibros(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                //total = total + RMDataCalculos9Extralibros.Count;

                RMDataCalculos9Extralibros = LiProcesoResultadoMensual(RMDataCalculos9Extralibros, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculos9Extralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensual> RMData = ProcesoResultadoMensual.ListarFromQueryCalculos1(_db, IdAgencia, anio, mes);
                total = RMData.Count;

                RMData = LiProcesoResultadoMensual(RMData, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensual item in RMData)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensual> RMDataCalculos = ProcesoResultadoMensual.ListarFromQueryCalculos2(_db, IdAgencia, anio, mes);

                total = total + RMDataCalculos.Count;

                RMDataCalculos = LiProcesoResultadoMensual(RMDataCalculos, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensual item in RMDataCalculos)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensual> RMDataCalculos2 = ProcesoResultadoMensual.ListarFromQueryCalculos3(_db, IdAgencia, anio, mes);

                total = total + RMDataCalculos2.Count;
                RMDataCalculos2 = LiProcesoResultadoMensual(RMDataCalculos2, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensual item in RMDataCalculos2)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }


                List<ProcesoResultadoMensual> RMDataCalculos4 = ProcesoResultadoMensual.ListarFromQueryCalculos4(_db, IdAgencia, anio, mes);

                total = total + RMDataCalculos4.Count;
                RMDataCalculos4 = LiProcesoResultadoMensual(RMDataCalculos4, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensual item in RMDataCalculos4)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensual> RMDataCalculos5 = ProcesoResultadoMensual.ListarFromQueryCalculos5(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                total = total + RMDataCalculos5.Count;
                RMDataCalculos5 = LiProcesoResultadoMensual(RMDataCalculos5, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensual item in RMDataCalculos5)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensual> RMDataCalculos6 = ProcesoResultadoMensual.ListarFromQueryCalculos6(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                total = total + RMDataCalculos6.Count;
                RMDataCalculos6 = LiProcesoResultadoMensual(RMDataCalculos6, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensual item in RMDataCalculos6)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensual> RMDataCalculos7 = ProcesoResultadoMensual.ListarFromQueryCalculos7(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                total = total + RMDataCalculos7.Count;
                RMDataCalculos7 = LiProcesoResultadoMensual(RMDataCalculos7, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensual item in RMDataCalculos7)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensual> RMDataCalculos8 = ProcesoResultadoMensual.ListarFromQueryCalculos8(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                total = total + RMDataCalculos8.Count;
                RMDataCalculos8 = LiProcesoResultadoMensual(RMDataCalculos8, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensual item in RMDataCalculos8)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                //Extralibros
                List<ProcesoResultadoMensualExtralibros> RMDataExtralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos1Extralibros(_db, IdAgencia, anio, mes);
                //total = RMDataExtralibros.Count;
                RMDataExtralibros = LiProcesoResultadoMensual(RMDataExtralibros, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensualExtralibros item in RMDataExtralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }


                List<ProcesoResultadoMensualExtralibros> RMDataCalculosExtralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos2Extralibros(_db, IdAgencia, anio, mes);

                //total = total + RMDataCalculosExtralibros.Count;
                RMDataCalculosExtralibros = LiProcesoResultadoMensual(RMDataCalculosExtralibros, IdAgencia, anio, mes);
                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculosExtralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }


                List<ProcesoResultadoMensualExtralibros> RMDataCalculos2Extralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos3Extralibros(_db, IdAgencia, anio, mes);

                //total = total + RMDataCalculos2Extralibros.Count;
                RMDataCalculos2Extralibros = LiProcesoResultadoMensual(RMDataCalculos2Extralibros, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculos2Extralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensualExtralibros> RMDataCalculos8Extralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos8Extralibros(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                RMDataCalculos8Extralibros = LiProcesoResultadoMensual(RMDataCalculos8Extralibros, IdAgencia, anio, mes);

                //total = total + RMDataCalculos8Extralibros.Count;

                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculos8Extralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensualExtralibros> RMDataCalculos4Extralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos4Extralibros(_db, IdAgencia, anio, mes);

                //total = total + RMDataCalculos4.Count;
                RMDataCalculos4Extralibros = LiProcesoResultadoMensual(RMDataCalculos4Extralibros, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculos4Extralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensualExtralibros> RMDataCalculos11Extralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos11Extralibros(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                //total = total + RMDataCalculos11Extralibros.Count;
                RMDataCalculos11Extralibros = LiProcesoResultadoMensual(RMDataCalculos11Extralibros, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculos11Extralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);

                }

                List<ProcesoResultadoMensualExtralibros> RMDataCalculos5Extralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos5Extralibros(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                //total = total + RMDataCalculos5.Count;
                RMDataCalculos5Extralibros = LiProcesoResultadoMensual(RMDataCalculos5Extralibros, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculos5Extralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensualExtralibros> RMDataCalculos6Extralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos6Extralibros(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                //total = total + RMDataCalculos6Extralibros.Count;
                RMDataCalculos6Extralibros = LiProcesoResultadoMensual(RMDataCalculos6Extralibros, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculos6Extralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                List<ProcesoResultadoMensualExtralibros> RMDataCalculos7Extralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos7Extralibros(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                //total = total + RMDataCalculos7Extralibros.Count;
                RMDataCalculos7Extralibros = LiProcesoResultadoMensual(RMDataCalculos7Extralibros, IdAgencia, anio, mes);

                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculos7Extralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                ProcesoResultadoMensual AbsorcionDeGastos = new ProcesoResultadoMensual();
                AbsorcionDeGastos.Id = 0;
                AbsorcionDeGastos.IdMarca = 0;
                AbsorcionDeGastos.IdAgencia = IdAgencia;
                AbsorcionDeGastos.Anio = anio;
                AbsorcionDeGastos.IdMes = mes;
                AbsorcionDeGastos.IdConcepto = 1060;
                AbsorcionDeGastos.Concepto = "ABSORCION DE GASTOS";
                AbsorcionDeGastos.Valor = 0;
                _db.Insert(14091, 999, AbsorcionDeGastos);

                ProcesoResultadoMensualExtralibros AbsorcionDeGastosEx = new ProcesoResultadoMensualExtralibros();
                AbsorcionDeGastosEx.Id = 0;
                AbsorcionDeGastosEx.IdMarca = 0;
                AbsorcionDeGastosEx.IdAgencia = IdAgencia;
                AbsorcionDeGastosEx.Anio = anio;
                AbsorcionDeGastosEx.IdMes = mes;
                AbsorcionDeGastosEx.IdConcepto = 1060;
                AbsorcionDeGastosEx.Concepto = "ABSORCION DE GASTOS";
                AbsorcionDeGastosEx.Valor = 0;
                _db.Insert(14091, 999, AbsorcionDeGastosEx);

                DataTable dtUV = _db.GetDataTable("SELECT * FROM [PREFIX]FINA.FNDRESMEN WHERE FIFNIDCIAU = " + IdAgencia + " AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1013) ORDER BY FIFNCPTD");
                if (dtUV.Rows.Count <= 0)
                {
                    ProcesoResultadoMensual NoFacturasEmitidasServicioEL = new ProcesoResultadoMensual();
                    NoFacturasEmitidasServicioEL.Id = 0;
                    NoFacturasEmitidasServicioEL.IdMarca = 0;
                    NoFacturasEmitidasServicioEL.IdAgencia = IdAgencia;
                    NoFacturasEmitidasServicioEL.Anio = anio;
                    NoFacturasEmitidasServicioEL.IdMes = mes;
                    NoFacturasEmitidasServicioEL.IdConcepto = 1013;
                    NoFacturasEmitidasServicioEL.Concepto = "AUTOS USADOS_UNIDADES SEMINUEVAS VENDIDAS";
                    NoFacturasEmitidasServicioEL.Valor = 0;
                    _db.Insert(14091, 999, NoFacturasEmitidasServicioEL);
                }

                DataTable dtUVEx = _db.GetDataTable("SELECT * FROM [PREFIX]FINA.FNDRSMENE WHERE FIFNIDCIAU = " + IdAgencia + " AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1013) ORDER BY FIFNCPTD");
                if (dtUVEx.Rows.Count <= 0)
                {
                    ProcesoResultadoMensualExtralibros NoFacturasEmitidasServicioEL = new ProcesoResultadoMensualExtralibros();
                    NoFacturasEmitidasServicioEL.Id = 0;
                    NoFacturasEmitidasServicioEL.IdMarca = 0;
                    NoFacturasEmitidasServicioEL.IdAgencia = IdAgencia;
                    NoFacturasEmitidasServicioEL.Anio = anio;
                    NoFacturasEmitidasServicioEL.IdMes = mes;
                    NoFacturasEmitidasServicioEL.IdConcepto = 1013;
                    NoFacturasEmitidasServicioEL.Concepto = "AUTOS USADOS_UNIDADES SEMINUEVAS VENDIDAS";
                    NoFacturasEmitidasServicioEL.Valor = 0;
                    _db.Insert(14091, 999, NoFacturasEmitidasServicioEL);
                }

                List<ProcesoResultadoMensualExtralibros> RMDataCalculos10Extralibros = ProcesoResultadoMensualExtralibros.ListarFromQueryCalculos10Extralibros(_db, IdAgencia, anio, mes).OrderBy(o => o.IdConcepto).ToList();

                //total = total + RMDataCalculos10Extralibros.Count;

                foreach (ProcesoResultadoMensualExtralibros item in RMDataCalculos10Extralibros)
                {
                    item.Id = 0;
                    _db.Insert(14091, 999, item);
                }

                #region No. Horas Facturadas

                DataTable dt = _db.GetDataTable("SELECT * FROM [PREFIX]FINA.FNDRSMENE WHERE FIFNIDCIAU = " + IdAgencia + " AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1031,1032,1033) ORDER BY FIFNCPTD");

                DateTime timenow = DateTime.Now;
                if (dt.Rows.Count == 0)
                {
                    if (diaPrimero)
                    {
                        //var query = @"delete FROM [PREFIX]FINA.FNDRESMEN
                        //WHERE FIFNIDCIAU = " + IdAgencia + @"
                        //AND FIFNYEAR = " + anio + @"
                        //AND FIFNMONTH = " + mes + @"
                        //AND FIFNCPTD IN (1031,1032,1033)";
                        //Console.WriteLine("Eliminando registros anteriores RM");
                        //var registrosEliminados = _db.SetQuery(query);

                        //var queryex = @"delete FROM [PREFIX]FINA.FNDRSMENE
                        //WHERE FIFNIDCIAU = " + IdAgencia + @"
                        //AND FIFNYEAR = " + anio + @"
                        //AND FIFNMONTH = " + mes + @"
                        //AND FIFNCPTD IN (1031,1032,1033)";
                        //Console.WriteLine("Eliminando registros anteriores RM Extralibros");
                        //var registrosEliminadosEx = _db.SetQuery(queryex);

                        SumaFacturas(_db, ref facTaller, ref facHYP, mes, anio, IdAgencia);
                        decimal TotalFacturasClienteTaller = facTaller.Sum() - facTaller[1];
                        decimal TotalFacturasHyPEx = facHYP.Sum() - facHYP[0];
                        decimal TotalFacturasHyP = facHYP.Sum() - facHYP[0];

                        //1031 - No. FACTURAS EMITIDAS SERVICIO
                        ProcesoResultadoMensualExtralibros NoFacturasEmitidasServicioEL = new ProcesoResultadoMensualExtralibros();
                        NoFacturasEmitidasServicioEL.Id = 0;
                        NoFacturasEmitidasServicioEL.IdMarca = 0;
                        NoFacturasEmitidasServicioEL.IdAgencia = IdAgencia;
                        NoFacturasEmitidasServicioEL.Anio = anio;
                        NoFacturasEmitidasServicioEL.IdMes = mes;
                        NoFacturasEmitidasServicioEL.IdConcepto = 1031;
                        NoFacturasEmitidasServicioEL.Concepto = "No. FACTURAS EMITIDAS SERVICIO";
                        NoFacturasEmitidasServicioEL.Valor = Convert.ToInt32(TotalFacturasClienteTaller);
                        _db.Insert(14091, 999, NoFacturasEmitidasServicioEL);

                        //1032 - No. FACTURAS EMITIDAS HyP
                        DataSet dsFACTHYPEx = ProcesoResultadoMensual.DevolverDatosFactHYP(_db, mes, anio, IdAgencia);
                        if (dsFACTHYPEx.Tables[0]?.Rows?.Count > 0)
                            TotalFacturasHyPEx = TotalFacturasHyPEx + Convert.ToInt32((dsFACTHYPEx.Tables[0]?.Rows?[0]["FACTURAS"].ToString().Trim() == "") ? "0" : dsFACTHYPEx.Tables[0]?.Rows?[0]["FACTURAS"].ToString().Trim());

                        ProcesoResultadoMensualExtralibros NoFacturasEmitidasHyP_Ex = new ProcesoResultadoMensualExtralibros();
                        NoFacturasEmitidasHyP_Ex.Id = 0;
                        NoFacturasEmitidasHyP_Ex.IdMarca = 0;
                        NoFacturasEmitidasHyP_Ex.IdAgencia = IdAgencia;
                        NoFacturasEmitidasHyP_Ex.Anio = anio;
                        NoFacturasEmitidasHyP_Ex.IdMes = mes;
                        NoFacturasEmitidasHyP_Ex.IdConcepto = 1032;
                        NoFacturasEmitidasHyP_Ex.Concepto = "No. FACTURAS EMITIDAS HyP";
                        NoFacturasEmitidasHyP_Ex.Valor = Convert.ToInt32(TotalFacturasHyPEx);
                        _db.Insert(14091, 999, NoFacturasEmitidasHyP_Ex);

                        //1033 - No. HORAS FACTURADAS SERVICIO
                        DataSet ds_Ex = ProcesoResultadoMensual.DevolverDatosOrdenadoHyP(_db, mes, anio, IdAgencia);
                        if (ds_Ex.Tables[0].Rows.Count > 0)
                        {
                            ProcesoResultadoMensualExtralibros NoHorasFacturadasServicio_Ex = new ProcesoResultadoMensualExtralibros();
                            NoHorasFacturadasServicio_Ex.Id = 0;
                            NoHorasFacturadasServicio_Ex.IdMarca = 0;
                            NoHorasFacturadasServicio_Ex.IdAgencia = IdAgencia;
                            NoHorasFacturadasServicio_Ex.Anio = anio;
                            NoHorasFacturadasServicio_Ex.IdMes = mes;
                            NoHorasFacturadasServicio_Ex.IdConcepto = 1033;
                            NoHorasFacturadasServicio_Ex.Concepto = "No. HORAS FACTURADAS SERVICIO";
                            NoHorasFacturadasServicio_Ex.Valor = Convert.ToInt32(ds_Ex.Tables[0].Rows[0]["UNIDADES_TIEMPO"]);
                            _db.Insert(14091, 999, NoHorasFacturadasServicio_Ex);
                        }
                        else
                        {
                            ProcesoResultadoMensualExtralibros NoHorasFacturadasServicio_Ex = new ProcesoResultadoMensualExtralibros();
                            NoHorasFacturadasServicio_Ex.Id = 0;
                            NoHorasFacturadasServicio_Ex.IdMarca = 0;
                            NoHorasFacturadasServicio_Ex.IdAgencia = IdAgencia;
                            NoHorasFacturadasServicio_Ex.Anio = anio;
                            NoHorasFacturadasServicio_Ex.IdMes = mes;
                            NoHorasFacturadasServicio_Ex.IdConcepto = 1033;
                            NoHorasFacturadasServicio_Ex.Concepto = "No. HORAS FACTURADAS SERVICIO";
                            NoHorasFacturadasServicio_Ex.Valor = 0;
                            _db.Insert(14091, 999, NoHorasFacturadasServicio_Ex);
                        }

                        //1031 - No. FACTURAS EMITIDAS SERVICIO
                        ProcesoResultadoMensual NoFacturasEmitidasServicio = new ProcesoResultadoMensual();
                        NoFacturasEmitidasServicio.Id = 0;
                        NoFacturasEmitidasServicio.IdMarca = 0;
                        NoFacturasEmitidasServicio.IdAgencia = IdAgencia;
                        NoFacturasEmitidasServicio.Anio = anio;
                        NoFacturasEmitidasServicio.IdMes = mes;
                        NoFacturasEmitidasServicio.IdConcepto = 1031;
                        NoFacturasEmitidasServicio.Concepto = "No. FACTURAS EMITIDAS SERVICIO";
                        NoFacturasEmitidasServicio.Valor = Convert.ToInt32(TotalFacturasClienteTaller);
                        _db.Insert(14091, 999, NoFacturasEmitidasServicio);


                        ResumenRMFacturasYHoras Resumen = new ResumenRMFacturasYHoras();
                        Resumen.IdAgencia = IdAgencia;
                        Resumen.Anio = anio;
                        Resumen.Mes = mes;
                        Resumen.IdConcepto = 1031;
                        Resumen.Valor = Convert.ToInt32(TotalFacturasClienteTaller);
                        _db.Insert(14091, 999, Resumen);

                        //DataSet dsFE = ProcesoResultadoMensual.DevolverDatosOrdenadoDetalle(_db, mes, anio, IdAgencia);
                        //for (int i = 0; i < dsFE.Tables[0].Rows.Count; i++)
                        //{
                        //    DetalleFacturas DetalleF = new DetalleFacturas();
                        //    DetalleF.IdAgencia = IdAgencia;
                        //    DetalleF.Anio = anio;
                        //    DetalleF.Mes = mes;
                        //    DetalleF.ConceptoVenta = Convert.ToInt32(dsFE.Tables[0].Rows[i]["FISEIDCNVN"]);
                        //    DetalleF.ConsecutivoVenta = Convert.ToInt32(dsFE.Tables[0].Rows[i]["FISEICNCVN"]);
                        //    DetalleF.TipoDePreorden = Convert.ToInt32(dsFE.Tables[0].Rows[i]["FISEIDTIPP"]);
                        //    DetalleF.IdAsesor = Convert.ToInt32(dsFE.Tables[0].Rows[i]["FISEIDASES"]);
                        //    DetalleF.FechaVenta = Convert.ToDateTime(dsFE.Tables[0].Rows[i]["FFSEFECCVN"]);
                        //    DetalleF.AnioMesVenta = Convert.ToInt32(dsFE.Tables[0].Rows[i]["FISEANMCVN"]);
                        //    DetalleF.MontoVenta = Convert.ToDecimal(dsFE.Tables[0].Rows[i]["FDSEMTOCVN"]);
                        //    DetalleF.CantidadVenta = Convert.ToDecimal(dsFE.Tables[0].Rows[i]["FDSECANCVN"]);
                        //    DetalleF.TipoDocumento = Convert.ToString(dsFE.Tables[0].Rows[i]["FSSETPDCVN"]);
                        //    DetalleF.DocumentoConceptoVenta = Convert.ToInt32(dsFE.Tables[0].Rows[i]["FISENDCCVN"]);
                        //    DetalleF.Folio = Convert.ToInt32(dsFE.Tables[0].Rows[0]["FISEFOLIO"]);
                        //    DetalleF.EConceptosDeVenta = Convert.ToString(dsFE.Tables[0].Rows[i]["FSSEDSCNVN"]);
                        //    DetalleF.TipoDeOrden = Convert.ToInt32(dsFE.Tables[0].Rows[i]["TIPOORDEN"]);
                        //    DetalleF.Tipo = Convert.ToInt32((dsFE.Tables[0].Rows[i]["FICAIDTINC"] == DBNull.Value) ? 0 : dsFE.Tables[0].Rows[i]["FICAIDTINC"]);
                        //    _db.Insert(14091, 999, DetalleF);

                        //}

                        //1032 - No. FACTURAS EMITIDAS HyP
                        DataSet dsFACTHYP = ProcesoResultadoMensual.DevolverDatosFactHYP(_db, mes, anio, IdAgencia);
                        if (dsFACTHYP.Tables[0]?.Rows?.Count > 0)
                            TotalFacturasHyP = TotalFacturasHyP + Convert.ToInt32((dsFACTHYP.Tables[0]?.Rows?[0]["FACTURAS"].ToString().Trim() == "") ? "0" : dsFACTHYP.Tables[0]?.Rows?[0]["FACTURAS"].ToString().Trim());

                        ProcesoResultadoMensual NoFacturasEmitidasHyP = new ProcesoResultadoMensual();
                        NoFacturasEmitidasHyP.Id = 0;
                        NoFacturasEmitidasHyP.IdMarca = 0;
                        NoFacturasEmitidasHyP.IdAgencia = IdAgencia;
                        NoFacturasEmitidasHyP.Anio = anio;
                        NoFacturasEmitidasHyP.IdMes = mes;
                        NoFacturasEmitidasHyP.IdConcepto = 1032;
                        NoFacturasEmitidasHyP.Concepto = "No. FACTURAS EMITIDAS HyP";
                        NoFacturasEmitidasHyP.Valor = Convert.ToInt32(TotalFacturasHyP);
                        _db.Insert(14091, 999, NoFacturasEmitidasHyP);

                        ResumenRMFacturasYHoras ResumenHyP = new ResumenRMFacturasYHoras();
                        ResumenHyP.IdAgencia = IdAgencia;
                        ResumenHyP.Anio = anio;
                        ResumenHyP.Mes = mes;
                        ResumenHyP.IdConcepto = 1032;
                        ResumenHyP.Valor = Convert.ToInt32(TotalFacturasHyP);
                        _db.Insert(14091, 999, ResumenHyP);


                        //DataSet dsDHP = ProcesoResultadoMensual.DevolverDatosFactHYPDetalle(_db, mes, anio, IdAgencia);
                        //for (int i = 0; i < dsDHP.Tables[0].Rows.Count; i++)
                        //{
                        //    DetalleFacturasHyP DetalleHP = new DetalleFacturasHyP();
                        //    DetalleHP.IdAgencia = IdAgencia;
                        //    DetalleHP.Anio = anio;
                        //    DetalleHP.Mes = mes;
                        //    DetalleHP.Siglas = Convert.ToString(dsDHP.Tables[0].Rows[i]["FSGESIGCIA"]);
                        //    DetalleHP.Agencia = Convert.ToString(dsDHP.Tables[0].Rows[i]["FSGERAZSOC"]);
                        //    DetalleHP.Agencia = DetalleHP.Agencia.TrimEnd();
                        //    DetalleHP.PrefijoInicial = Convert.ToString(dsDHP.Tables[0].Rows[i]["FCCAPREIN"]);
                        //    DetalleHP.FolioInicial = Convert.ToInt32(dsDHP.Tables[0].Rows[i]["FICAFOLIN"]);
                        //    DetalleHP.Fecha = Convert.ToDateTime(dsDHP.Tables[0].Rows[i]["FFCAFECHA"]);
                        //    DetalleHP.Facnot = Convert.ToString(dsDHP.Tables[0].Rows[i]["FACTONOTA"]);
                        //    DetalleHP.FacturaOriginal_Folio = Convert.ToInt32(dsDHP.Tables[0].Rows[i]["FISEIDTIPP"]);
                        //    DetalleHP.TipoDePreorden = Convert.ToInt32(dsDHP.Tables[0].Rows[i][9]);
                        //    DetalleHP.Clave = Convert.ToString(dsDHP.Tables[0].Rows[i]["FSCACLAVE"]);
                        //    _db.Insert(14091, 999, DetalleHP);

                        //}

                        //1033 - No. HORAS FACTURADAS SERVICIO
                        DataSet ds = ProcesoResultadoMensual.DevolverDatosOrdenadoHyP(_db, mes, anio, IdAgencia);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            ProcesoResultadoMensual NoHorasFacturadasServicio = new ProcesoResultadoMensual();
                            NoHorasFacturadasServicio.Id = 0;
                            NoHorasFacturadasServicio.IdMarca = 0;
                            NoHorasFacturadasServicio.IdAgencia = IdAgencia;
                            NoHorasFacturadasServicio.Anio = anio;
                            NoHorasFacturadasServicio.IdMes = mes;
                            NoHorasFacturadasServicio.IdConcepto = 1033;
                            NoHorasFacturadasServicio.Concepto = "No. HORAS FACTURADAS SERVICIO";
                            NoHorasFacturadasServicio.Valor = Convert.ToInt32(ds.Tables[0].Rows[0]["UNIDADES_TIEMPO"]);
                            _db.Insert(14091, 999, NoHorasFacturadasServicio);

                            if (diaPrimero)
                            {
                                ResumenRMFacturasYHoras ResumenHF = new ResumenRMFacturasYHoras();
                                ResumenHF.IdAgencia = IdAgencia;
                                ResumenHF.Anio = anio;
                                ResumenHF.Mes = mes;
                                ResumenHF.IdConcepto = 1033;
                                ResumenHF.Valor = Convert.ToInt32(ds.Tables[0].Rows[0]["UNIDADES_TIEMPO"]);
                                _db.Insert(14091, 999, ResumenHF);

                                //DataSet dsHF = ProcesoResultadoMensual.DevolverDatosOrdenadoHyPDetalle(_db, mes, anio, IdAgencia);
                                //for (int i = 0; i < dsHF.Tables[0].Rows.Count; i++)
                                //{
                                //    DetalleHorasFacturadas DetalleHoras = new DetalleHorasFacturadas();
                                //    DetalleHoras.IdAgencia = IdAgencia;
                                //    DetalleHoras.Anio = anio;
                                //    DetalleHoras.Mes = mes;
                                //    DetalleHoras.Siglas = Convert.ToString(dsHF.Tables[0].Rows[i]["SIGLAS"]);
                                //    DetalleHoras.FacturaOriginal_Folio = Convert.ToInt32(dsHF.Tables[0].Rows[i]["ID_FACTURA"]);
                                //    DetalleHoras.IdOperacion = Convert.ToInt32(dsHF.Tables[0].Rows[i]["ID_OPERACION"]);
                                //    DetalleHoras.NumeroUnidades = Convert.ToDecimal(dsHF.Tables[0].Rows[i]["UNIDADES_TIEMPO"]);
                                //    _db.Insert(14091, 999, DetalleHoras);

                                //}
                            }

                        }
                        else
                        {
                            ProcesoResultadoMensual NoHorasFacturadasServicio = new ProcesoResultadoMensual();
                            NoHorasFacturadasServicio.Id = 0;
                            NoHorasFacturadasServicio.IdMarca = 0;
                            NoHorasFacturadasServicio.IdAgencia = IdAgencia;
                            NoHorasFacturadasServicio.Anio = anio;
                            NoHorasFacturadasServicio.IdMes = mes;
                            NoHorasFacturadasServicio.IdConcepto = 1033;
                            NoHorasFacturadasServicio.Concepto = "No. HORAS FACTURADAS SERVICIO";
                            NoHorasFacturadasServicio.Valor = 0;
                            _db.Insert(14091, 999, NoHorasFacturadasServicio);
                        }
                    }
                }

                #endregion

                RecalculaUtilidadBrutaXUnidad(IdAgencia, anio, mes);
                RecalculaGastosAdministacionYUtilidadNetaDepartamentos(IdAgencia, anio, mes);
                RecalculaIngresoTotalEIngresoGerenciaDeNegocios(IdAgencia, anio, mes);
                if (anio == 2024 && mes == 4)
                    RecalculaGastosAdministacionXDepartamento(IdAgencia, anio, mes);
                else if (anio == 2024 && mes >= 5)
                {
                    RecalculaGastosAdministacionYOtrosGastosProductos(IdAgencia, anio, mes);
                    RecalculaUtilidadDeOperacion(IdAgencia, anio, mes);
                    RecalculaGastosAdministacionXDepartamento(IdAgencia, anio, mes);
                    RecalculaUNAFyUFAIyUN(IdAgencia, anio, mes);
                }

                DateTime dateFin = DateTime.Now;
                TimeSpan TimeDif = new TimeSpan();
                TimeDif = dateFin - dateIni;
                Console.WriteLine("[FINALIZA RM][TotalTime=" + TimeDif + "][TotalInsert=" + total + "]Proceso Terminado para la agencia ID: " + IdAgencia);
            }
            catch (Exception e)
            {
                EnviarCorreoError(_db, e.Message, IdAgencia, anio, mes);
                Console.WriteLine("[ERROR][EjecutaParaAgencia][Agencia=" + IdAgencia + "] Error: " + e);
            }
        }

        public void EnviarCorreoError(DB2Database aDB, string Error, int IdAgencia, int Año, int Mes)
        {
            DateTime dt = DateTime.Now;
            SendEmail sendEmailXDoc = new SendEmail();
            sendEmailXDoc.AgregaDestinatarios("jvargasr@grupoautofin.com");

            sendEmailXDoc.SenderMail = aDB.CA_GE_CORREO_DE_NOTIFICACIONES.ToLower().Trim();
            sendEmailXDoc.Password = aDB.CA_GE_CONTRASENIA_CORREO_DE_NOTIFICACIONES;
            sendEmailXDoc.User = sendEmailXDoc.SenderMail;

            StringBuilder strContenidoXDoc = new StringBuilder();
            strContenidoXDoc.Append("\r\n");
            strContenidoXDoc.Append(String.Format("Se informa que el proceso : {0} ", "ProcesoLlenadoDeTablaArchivoBase"));
            strContenidoXDoc.Append(String.Format("Año : {0} ", dt.Year));
            strContenidoXDoc.Append(String.Format("del Mes : {0} ", dt.Month - 1));
            strContenidoXDoc.Append(String.Format("Agencia : {0} ", IdAgencia));
            strContenidoXDoc.Append("\r\n\r\n");
            strContenidoXDoc.Append(String.Format("Tubo un error: "));
            strContenidoXDoc.Append("\r\n\r\n");
            strContenidoXDoc.Append("***************** Error *******************.\r\n\r\n");
            strContenidoXDoc.Append(String.Format("Error : {0} ", Error));
            sendEmailXDoc.EnviaCorreoWebMail("[ERROR]Informe Procesos de Finanzas.", strContenidoXDoc.ToString(), false, System.Net.Mail.MailPriority.High);
        }

        public List<ProcesoResultadoMensual> LiProcesoResultadoMensual(List<ProcesoResultadoMensual> liRM, int Agencia, int Año, int Mes)
        {
            foreach (ProcesoResultadoMensual RM in liRM)
            {
                List<int> departamentos = new List<int>() { 5 };
                switch (RM.IdConcepto)
                {
                    case 1005:
                        departamentos.Add(22);
                        break;
                    case 1011:
                        departamentos.Add(26);
                        break;
                    case 1017:
                        departamentos.Add(27);
                        break;
                    case 1023:
                        departamentos.Add(28);
                        break;
                    case 1027:
                        departamentos.Add(29);
                        break;
                    case 1040:
                        departamentos.Add(30);
                        break;
                    case 1044:
                        departamentos.Add(31);
                        break;
                }

                List<AjusteImporteConceptos> liAjustes = new List<AjusteImporteConceptos>();
                foreach (int depart in departamentos)
                {
                    if (depart == 5)
                        liAjustes = AjusteImporteConceptos.ListarPorAnioMesAgenciaReporteConcepto(_db, Agencia, Año, Mes, RM.IdConcepto, depart, false);
                    else
                        liAjustes.AddRange(AjusteImporteConceptos.ListarPorAnioMesAgenciaYReporte(_db, Agencia, Año, Mes, depart, true));
                }

                foreach (AjusteImporteConceptos Ajuste in liAjustes)
                {
                    if ((Ajuste.Id != 45) || (Ajuste.Id != 46))
                    {
                        ConceptosContables Concepto = liConceptos.Where(x => x.Id == Ajuste.IdConcepto).SingleOrDefault();

                        //if (Ajuste.IdPlantilla == 1)
                        //{
                        //    if (Ajuste.Sesuma)
                        //        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000) * -1;
                        //    else
                        //        RM.Valor -= Convert.ToInt32(Ajuste.ImporteAjustes * 1000) * -1;
                        //}
                        //else
                        //{
                        //    if (Concepto.Id == 1049)
                        //        RM.Valor -= Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        //    else
                        //    {
                        //        if (Ajuste.Sesuma)
                        //            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        //        else
                        //            RM.Valor -= Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        //    }
                        //}

                        if (Concepto.Id == RM.IdConcepto)
                            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        else if (Ajuste.IdReporte == 22 && RM.IdConcepto == 1005)
                            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        else if (Ajuste.IdReporte == 26 && RM.IdConcepto == 1011)
                            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        else if (Ajuste.IdReporte == 27 && RM.IdConcepto == 1017)
                            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        else if (Ajuste.IdReporte == 28 && RM.IdConcepto == 1023)
                            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        else if (Ajuste.IdReporte == 29 && RM.IdConcepto == 1027)
                            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        else if (Ajuste.IdReporte == 30 && RM.IdConcepto == 1040)
                            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                        else if (Ajuste.IdReporte == 31 && RM.IdConcepto == 1044)
                            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    }
                }

            }

            return liRM;
        }

        public List<ProcesoResultadoMensualExtralibros> LiProcesoResultadoMensual(List<ProcesoResultadoMensualExtralibros> liRM, int Agencia, int Año, int Mes)
        {
            foreach (ProcesoResultadoMensualExtralibros RM in liRM)
            {
                List<int> departamentos = new List<int>() { 5 };
                switch (RM.IdConcepto)
                {
                    case 1005:
                        departamentos.Add(22);
                        break;
                    case 1011:
                        departamentos.Add(26);
                        break;
                    case 1017:
                        departamentos.Add(27);
                        break;
                    case 1023:
                        departamentos.Add(28);
                        break;
                    case 1027:
                        departamentos.Add(29);
                        break;
                    case 1040:
                        departamentos.Add(30);
                        break;
                    case 1044:
                        departamentos.Add(31);
                        break;
                }

                List<AjusteImporteConceptos> liAjustes = new List<AjusteImporteConceptos>();
                foreach (int depart in departamentos)
                {
                    if (depart == 5)
                        liAjustes = AjusteImporteConceptos.ListarPorAnioMesAgenciaReporteConcepto(_db, Agencia, Año, Mes, RM.IdConcepto, depart, false);
                    else
                        liAjustes.AddRange(AjusteImporteConceptos.ListarPorAnioMesAgenciaYReporte(_db, Agencia, Año, Mes, depart, true));
                }

                foreach (AjusteImporteConceptos Ajuste in liAjustes)
                {
                    ConceptosContables Concepto = liConceptos.Where(x => x.Id == Ajuste.IdConcepto).SingleOrDefault();

                    //if (Ajuste.IdPlantilla == 1)
                    //{
                    //    if (Ajuste.Sesuma)
                    //        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000) * -1;
                    //    else
                    //        RM.Valor -= Convert.ToInt32(Ajuste.ImporteAjustes * 1000) * -1;
                    //}
                    //else
                    //{
                    //    if (Concepto.Id == 1049)
                    //        RM.Valor -= Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    //    else
                    //    {
                    //        if (Ajuste.Sesuma)
                    //            RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    //        else
                    //            RM.Valor -= Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    //    }
                    //}

                    if (Concepto.Id == RM.IdConcepto)
                        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    else if (Ajuste.IdReporte == 22 && RM.IdConcepto == 1005)
                        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    else if (Ajuste.IdReporte == 26 && RM.IdConcepto == 1011)
                        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    else if (Ajuste.IdReporte == 27 && RM.IdConcepto == 1017)
                        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    else if (Ajuste.IdReporte == 28 && RM.IdConcepto == 1023)
                        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    else if (Ajuste.IdReporte == 29 && RM.IdConcepto == 1027)
                        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    else if (Ajuste.IdReporte == 30 && RM.IdConcepto == 1040)
                        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                    else if (Ajuste.IdReporte == 31 && RM.IdConcepto == 1044)
                        RM.Valor += Convert.ToInt32(Ajuste.ImporteAjustes * 1000);
                }

            }

            return liRM;
        }

        public void RecalculaUtilidadBrutaXUnidad(int idAgencia, int anio, int mes)
        {
            query = "SELECT * FROM [PREFIX]FINA . FNDAGSUC WHERE FIFNIDCIAU = " + idAgencia + " AND FIFNSTATUS = 1 ORDER BY FIFNIDCIAS";

            DataTable dtSucursales = _db.GetDataTable(query);

            if (dtSucursales.Rows.Count != 0)
            {
                string idsSuc = "";

                foreach (DataRow drSuc in dtSucursales.Rows)
                {
                    idsSuc += drSuc["FIFNIDCIAS"].ToString() + ",";
                }

                idsSuc = idsSuc.Remove(idsSuc.Length - 1, 1);

                //V1
                //1001_AUTOS NUEVOS_UNIDADES NUEVAS VENDIDAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1001)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                DataTable dtUnidades = _db.GetDataTable(query);

                double unidades = 0.00;

                if (dtUnidades.Rows.Count != 0)
                    unidades = Convert.ToDouble(dtUnidades.Rows[0]["VALOR"]);

                //1004_AUTOS NUEVOS_UTILIDAD BRUTA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1004)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                DataTable dtUB = _db.GetDataTable(query);

                double utilidadBruta = 0.00;

                if (dtUB.Rows.Count != 0)
                    utilidadBruta = Convert.ToDouble(dtUB.Rows[0]["VALOR"]) / 1000;

                double ubXUnidad = 0.00;

                //1003_AUTOS NUEVOS_PROMEDIO UTILIDAD POR UNIDAD
                if (utilidadBruta != 0.00 && unidades != 0.00)
                {
                    ubXUnidad = utilidadBruta / unidades;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1003";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + ubXUnidad + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1003";

                    _db.SetQuery(query);
                }

                //V2
                //1001_AUTOS NUEVOS_UNIDADES NUEVAS VENDIDAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1001)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
                                
                dtUnidades = _db.GetDataTable(query);

                if (dtUnidades.Rows.Count != 0)
                    unidades = Convert.ToDouble(dtUnidades.Rows[0]["VALOR"]);

                //1004_AUTOS NUEVOS_UTILIDAD BRUTA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1004)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtUB = _db.GetDataTable(query);

                if (dtUB.Rows.Count != 0)
                    utilidadBruta = Convert.ToDouble(dtUB.Rows[0]["VALOR"]) / 1000;

                //1003_AUTOS NUEVOS_PROMEDIO UTILIDAD POR UNIDAD
                if (utilidadBruta != 0.00 && unidades != 0.00)
                {
                    ubXUnidad = utilidadBruta / unidades;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1003";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + ubXUnidad + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1003";

                    _db.SetQuery(query);
                }




                //V1
                //1013_UNIDADES SEMINUEVAS VENDIDAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1013)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtUnidades = _db.GetDataTable(query);

                unidades = 0.00;

                if (dtUnidades.Rows.Count != 0)
                    unidades = Convert.ToDouble(dtUnidades.Rows[0]["VALOR"]);

                //1016_UTILIDAD BRUTA DEPARTAMENTAL AU
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1016)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtUB = _db.GetDataTable(query);

                utilidadBruta = 0.00;

                if (dtUB.Rows.Count != 0)
                    utilidadBruta = Convert.ToDouble(dtUB.Rows[0]["VALOR"]) / 1000;

                ubXUnidad = 0.00;

                //1015_PROMEDIO DE UTILIDAD BRUTA POR UNIDAD
                if (utilidadBruta != 0.00 && unidades != 0.00)
                {
                    ubXUnidad = utilidadBruta / unidades;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1015";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + ubXUnidad + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1015";

                    _db.SetQuery(query);
                }

                //V2
                //1013_UNIDADES SEMINUEVAS VENDIDAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1013)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtUnidades = _db.GetDataTable(query);

                if (dtUnidades.Rows.Count != 0)
                    unidades = Convert.ToDouble(dtUnidades.Rows[0]["VALOR"]);

                //1016_UTILIDAD BRUTA DEPARTAMENTAL AU
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1016)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtUB = _db.GetDataTable(query);

                if (dtUB.Rows.Count != 0)
                    utilidadBruta = Convert.ToDouble(dtUB.Rows[0]["VALOR"]) / 1000;

                //1015_PROMEDIO DE UTILIDAD BRUTA POR UNIDAD
                if (utilidadBruta != 0.00 && unidades != 0.00)
                {
                    ubXUnidad = utilidadBruta / unidades;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1015";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + ubXUnidad + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1015";

                    _db.SetQuery(query);
                }
            }
        }

        public void RecalculaGastosAdministacionYUtilidadNetaDepartamentos(int idAgencia, int anio, int mes)
        {
            query = "SELECT * FROM [PREFIX]FINA . FNDAGSUC WHERE FIFNIDCIAU = " + idAgencia + " AND FIFNSTATUS = 1 ORDER BY FIFNIDCIAS";

            DataTable dtSucursales = _db.GetDataTable(query);

            if (dtSucursales.Rows.Count != 0)
            {
                string idsSuc = "";

                foreach (DataRow drSuc in dtSucursales.Rows)
                {
                    idsSuc += drSuc["FIFNIDCIAS"].ToString() + ",";
                }

                idsSuc = idsSuc.Remove(idsSuc.Length - 1, 1);

                //V1
                //1000_INGRESOS TOTALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1000)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                DataTable dtInfo = _db.GetDataTable(query);

                double ingresoTotal = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoTotal = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1002_INGRESOS POR VENTA DE UNIDADES NUEVAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1002)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                double ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1008_UTILIDAD NETA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1008)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                double utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1085_FINANCIAMIENTO NETO AUTOS NUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1085)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                double finanNeto = 0.00;

                if (dtInfo.Rows.Count != 0)
                    finanNeto = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1044_GASTOS ADMINISTRATIVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                double gastosAdmin = 0.00;

                if (dtInfo.Rows.Count != 0)
                    gastosAdmin = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                double gastosAdminXDep = 0.00;
                double utilidadNetaXDep = 0.00;

                //V1
                //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";

                    _db.SetQuery(query);
                }

                //V1
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1009)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1012)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                finanNeto = 0.00;

                //V1
                //1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";

                    _db.SetQuery(query);
                }

                //V1
                //1014_INGRESOS VENTA DE UNIDADES SEMINUEVAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1014)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1020_UTILIDAD NETA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1020)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1086)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                finanNeto = 0.00;

                if (dtInfo.Rows.Count != 0)
                    finanNeto = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";

                    _db.SetQuery(query);
                }

                //V1
                //1021_INGRESOS POR VENTA DE SERVICIO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1021)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1024_UTILIDAD NETA SERVICIO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1024)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                finanNeto = 0.00;

                //V1
                //1091_GASTOS DE ADMINISTRACION SERVICIO
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";

                    _db.SetQuery(query);
                }

                //V1
                //1025_INGRESOS POR HOJALATERIA Y PINTURA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1025)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1028)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                finanNeto = 0.00;

                //V1
                //1092_GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";

                    _db.SetQuery(query);
                }

                //V1
                //1034_INGRESOS POR VENTA DE REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1034)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1041_UTILIDAD NETA REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1041)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1087_FINANCIAMIENTO NETO REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1087)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                finanNeto = 0.00;

                if (dtInfo.Rows.Count != 0)
                    finanNeto = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1093_GASTOS DE ADMINISTRACION REFACCIONES
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";

                    _db.SetQuery(query);
                }

                //V1
                //1061_GASTOS TOTALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1061)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                double gastoTotal = 0.00;

                if (dtInfo.Rows.Count != 0)
                    gastoTotal = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1062_UTILIDAD BRUTA DE SERVICIO Y 1063_UTILIDAD BRUTA DE REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1062,1063)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH ";

                dtInfo = _db.GetDataTable(query);

                double utilidadBruta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadBruta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1004_UTILIDAD BRUTA DEPARTAMENTAL AN Y 1016_UTILIDAD BRUTA DEPARTAMENTAL AU Y 1022_UTILIDAD BRUTA DEPARTAMENTAL SE Y 1026_UTILIDAD BRUTA DEPARTAMENTAL HYP
                //1039_UTILIDAD BRUTA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1004,1016,1022,1026,1039)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH ";

                dtInfo = _db.GetDataTable(query);

                double utilidadBrutaTodas = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadBrutaTodas = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1064_% DE ABSORCIÓN
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    double porcenAbsorcion = (utilidadBruta / gastoTotal) * 100;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1064";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + porcenAbsorcion + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1064";

                    _db.SetQuery(query);
                }

                //V1
                //1065_% DE GASTOS SOBRE UTILIDAD BRUTA TOTAL
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    double porcenGastos = (gastoTotal / utilidadBrutaTodas) * 100;

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1065";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + porcenGastos + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1065";

                    _db.SetQuery(query);
                }



                //V2
                //1000_INGRESOS TOTALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1000)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoTotal = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoTotal = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1002_INGRESOS POR VENTA DE UNIDADES NUEVAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1002)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1008_UTILIDAD NETA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1008)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1085_FINANCIAMIENTO NETO AUTOS NUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1085)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                finanNeto = 0.00;

                if (dtInfo.Rows.Count != 0)
                    finanNeto = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1044_GASTOS ADMINISTRATIVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                gastosAdmin = 0.00;

                if (dtInfo.Rows.Count != 0)
                    gastosAdmin = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                gastosAdminXDep = 0.00;

                //V2
                //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";

                    _db.SetQuery(query);
                }

                //V2
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1009)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1012)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                finanNeto = 0.00;

                //V2
                //1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";

                    _db.SetQuery(query);
                }

                //V2
                //1014_INGRESOS VENTA DE UNIDADES SEMINUEVAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1014)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1020_UTILIDAD NETA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1020)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1086)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                finanNeto = 0.00;

                if (dtInfo.Rows.Count != 0)
                    finanNeto = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";

                    _db.SetQuery(query);
                }

                //V2
                //1021_INGRESOS POR VENTA DE SERVICIO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1021)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1024_UTILIDAD NETA SERVICIO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1024)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                finanNeto = 0.00;

                //V2
                //1091_GASTOS DE ADMINISTRACION SERVICIO
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";

                    _db.SetQuery(query);
                }

                //V2
                //1025_INGRESOS POR HOJALATERIA Y PINTURA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1025)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1028)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                finanNeto = 0.00;

                //V2
                //1092_GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";

                    _db.SetQuery(query);

                    //V2
                    //2027_UTILIDAD NETA DEPARTAMENTAL SERVICIO
                    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                        "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (2027)\r\n" +
                        "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                    dtInfo = _db.GetDataTable(query);

                    double utilidadNetaServicio = 0.00;

                    if (dtInfo.Rows.Count != 0)
                        utilidadNetaServicio = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                    //V2
                    //2028_UTILIDAD NETA DEPARTAMENTAL HOJALATERIA Y PINTURA
                    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                        "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (2028)\r\n" +
                        "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                    dtInfo = _db.GetDataTable(query);

                    double utilidadNetaHyP = 0.00;

                    if (dtInfo.Rows.Count != 0)
                        utilidadNetaHyP = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                    //V2
                    //1029_UTILIDAD NETA SERVICIOS ADICIONALES
                    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                        "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1029)\r\n" +
                        "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                    dtInfo = _db.GetDataTable(query);

                    double utilidadNetaServAdic = 0.00;

                    if (dtInfo.Rows.Count != 0)
                        utilidadNetaServAdic = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                    double utilidadServicioYHyP = utilidadNetaServicio + utilidadNetaHyP + utilidadNetaServAdic;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1030";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadServicioYHyP + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1030";

                    _db.SetQuery(query);
                }

                //V2
                //1034_INGRESOS POR VENTA DE REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1034)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1041_UTILIDAD NETA REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1041)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                utilidadNeta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadNeta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1087_FINANCIAMIENTO NETO REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1087)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                finanNeto = 0.00;

                if (dtInfo.Rows.Count != 0)
                    finanNeto = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1093_GASTOS DE ADMINISTRACION REFACCIONES
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    gastosAdminXDep = ((ingresoDep / ingresoTotal) * gastosAdmin);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNeta + finanNeto - gastosAdminXDep;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";

                    _db.SetQuery(query);
                }

                //V2
                //1061_GASTOS TOTALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1061)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                gastoTotal = 0.00;

                if (dtInfo.Rows.Count != 0)
                    gastoTotal = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1062_UTILIDAD BRUTA DE SERVICIO Y 1063_UTILIDAD BRUTA DE REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1062,1063)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH";

                dtInfo = _db.GetDataTable(query);

                utilidadBruta = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadBruta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1004_UTILIDAD BRUTA DEPARTAMENTAL AN Y 1016_UTILIDAD BRUTA DEPARTAMENTAL AU Y 1022_UTILIDAD BRUTA DEPARTAMENTAL SE Y 1026_UTILIDAD BRUTA DEPARTAMENTAL HYP
                //1039_UTILIDAD BRUTA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1004,1016,1022,1026,1039)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH";

                dtInfo = _db.GetDataTable(query);

                utilidadBrutaTodas = 0.00;

                if (dtInfo.Rows.Count != 0)
                    utilidadBrutaTodas = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V2
                //1064_% DE ABSORCIÓN
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    double porcenAbsorcion = (utilidadBruta / gastoTotal) * 100;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1064";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + porcenAbsorcion + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1064";

                    _db.SetQuery(query);
                }

                //V2
                //1065_% DE GASTOS SOBRE UTILIDAD BRUTA TOTAL
                if (ingresoTotal != 0.00 && ingresoDep != 0.00)
                {
                    double porcenGastos = (gastoTotal / utilidadBrutaTodas) * 100;

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1065";

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + porcenGastos + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1065";

                    _db.SetQuery(query);
                }
            }
        }

        public void RecalculaIngresoTotalEIngresoGerenciaDeNegocios(int idAgencia, int anio, int mes)
        {
            if (idAgencia == 275) //275_PASION MOTORS HIDALGO SA DE CV
            {
                double ajuste = 0.00;

                query = "SELECT FIFNIMPAJU VALOR " +
                    "FROM [PREFIX]FINA.FNCAJICON " +
                    "WHERE FIFNIDCIAU = " + idAgencia + " AND FIFNIDPLAJ = 33 AND FIFNCPTD = 1010 " +
                    "AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes;

                DataTable dtInfo = _db.GetDataTable(query);

                if (dtInfo.Rows.Count != 0)
                    ajuste = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1000_INGRESOS TOTALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1000)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                double ingresoTotal = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoTotal = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1000_INGRESOS TOTALES
                if (ingresoTotal != 0.00 && ajuste != 0.00)
                {
                    ingresoTotal = (ingresoTotal + (ajuste * 1000));

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + ingresoTotal + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1000";

                    _db.SetQuery(query);
                }

                double ingresoDep = 0.00;

                //V1
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1009)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                if (ingresoDep != 0.00 && ajuste != 0.00)
                {
                    ingresoDep = (ingresoDep + (ajuste * 1000));

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + ingresoDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1009";

                    _db.SetQuery(query);
                }



                //V2
                //1000_INGRESOS TOTALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1000)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoTotal = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoTotal = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1000_INGRESOS TOTALES
                if (ingresoTotal != 0.00 && ajuste != 0.00)
                {
                    ingresoTotal = (ingresoTotal + (ajuste * 1000));

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + ingresoTotal + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1000";

                    _db.SetQuery(query);
                }

                //V2
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1009)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                if (ingresoDep != 0.00 && ajuste != 0.00)
                {
                    ingresoDep = (ingresoDep + (ajuste * 1000));

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + ingresoDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1009";

                    _db.SetQuery(query);
                }
            }

            if (idAgencia == 275) //275_PASION MOTORS HIDALGO SA DE CV
            {
                double ajuste = 0.00;

                query = "SELECT FIFNIMPAJU VALOR " +
                    "FROM [PREFIX]FINA.FNCAJICON " +
                    "WHERE FIFNIDCIAU = " + idAgencia + " AND FIFNIDPLAJ = 33 AND FIFNCPTD = 1010 " +
                    "AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes;

                DataTable dtInfo = _db.GetDataTable(query);

                if (dtInfo.Rows.Count != 0)
                    ajuste = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1000_INGRESOS TOTALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1000)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                double ingresoTotal = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoTotal = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1000_INGRESOS TOTALES
                if (ingresoTotal != 0.00 && ajuste != 0.00)
                {
                    ingresoTotal = (ingresoTotal + (ajuste * 1000));

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + ingresoTotal + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1000";

                    _db.SetQuery(query);
                }

                double ingresoDep = 0.00;

                //V1
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1009)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                if (ingresoDep != 0.00 && ajuste != 0.00)
                {
                    ingresoDep = (ingresoDep + (ajuste * 1000));

                    query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
                        "SET FIFNVALUE = " + ingresoDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1009";

                    _db.SetQuery(query);
                }



                //V2
                //1000_INGRESOS TOTALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1000)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoTotal = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoTotal = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1000_INGRESOS TOTALES
                if (ingresoTotal != 0.00 && ajuste != 0.00)
                {
                    ingresoTotal = (ingresoTotal + (ajuste * 1000));

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + ingresoTotal + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1000";

                    _db.SetQuery(query);
                }

                //V2
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1009)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";

                dtInfo = _db.GetDataTable(query);

                ingresoDep = 0.00;

                if (dtInfo.Rows.Count != 0)
                    ingresoDep = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //V1
                //1009_INGRESOS GERENCIA DE NEGOCIOS
                if (ingresoDep != 0.00 && ajuste != 0.00)
                {
                    ingresoDep = (ingresoDep + (ajuste * 1000));

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + ingresoDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1009";

                    _db.SetQuery(query);
                }
            }
        }

        public void RecalculaGastosAdministacionXDepartamento(int idAgencia, int anio, int mes)
        {
            string idsSuc = "";

            query = "SELECT * FROM [PREFIX]FINA . FNDAGSUC WHERE FIFNIDCIAU = " + idAgencia + " AND FIFNSTATUS = 1 ORDER BY FIFNIDCIAS";

            DataTable dtSucursales = _db.GetDataTable(query);

            if (dtSucursales.Rows.Count != 0)
            {
                foreach (DataRow drSuc in dtSucursales.Rows)
                {
                    idsSuc += drSuc["FIFNIDCIAS"].ToString() + ",";
                }

                idsSuc = idsSuc.Remove(idsSuc.Length - 1, 1);
            }

            #region V1

            //#region 1044_GASTOS ADMINISTRATIVOS

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1044_GASTOS ADMINISTRATIVOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //        "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
            //        "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1044_GASTOS ADMINISTRATIVOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //        "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
            //        "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            
            //DataTable dtInfo = _db.GetDataTable(query);

            //double gastosAdmin = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    gastosAdmin = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            //#endregion

            //#region 1008_UTILIDAD NETA DEPARTAMENTAL AUTOS NUEVOS

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1008_UTILIDAD NETA DEPARTAMENTAL AUTOS NUEVOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1008)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1008_UTILIDAD NETA DEPARTAMENTAL AUTOS NUEVOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1008)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}

            //dtInfo = _db.GetDataTable(query);

            //double utilidadNetaAN = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    utilidadNetaAN = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            //#endregion

            //#region 1085_FINANCIAMIENTO NETO AUTOS NUEVOS

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1085_FINANCIAMIENTO NETO AUTOS NUEVOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1085)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1085_FINANCIAMIENTO NETO AUTOS NUEVOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1085)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}

            //dtInfo = _db.GetDataTable(query);

            //double finanNetoAN = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    finanNetoAN = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            //#endregion

            //#region 1012_UTILIDAD NETA GERENCIA DE NEGOCIOS

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1012)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1012)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}

            //dtInfo = _db.GetDataTable(query);

            //double utilidadNetaGN = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    utilidadNetaGN = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            //double finanNetoGN = 0.00;

            //#endregion

            //#region 1020_UTILIDAD NETA DEPARTAMENTAL

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1020_UTILIDAD NETA DEPARTAMENTAL
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1020)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1020_UTILIDAD NETA DEPARTAMENTAL
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1020)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}

            //dtInfo = _db.GetDataTable(query);

            //double utilidadNetaAU = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    utilidadNetaAU = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            //#endregion

            //#region 1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1086)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1086)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}

            //dtInfo = _db.GetDataTable(query);

            //double finanNetoAU = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    finanNetoAU = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            //#endregion

            //#region 1024_UTILIDAD NETA SERVICIO

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1024_UTILIDAD NETA SERVICIO
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1024)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1024_UTILIDAD NETA SERVICIO
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1024)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}

            //dtInfo = _db.GetDataTable(query);

            //double utilidadNetaSE = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    utilidadNetaSE = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);
                        
            //double finanNetoSE = 0.00;

            //#endregion

            //#region 1028_UTILIDAD NETA HOJALATERIA Y PINTURA

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1028)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1028)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}

            //dtInfo = _db.GetDataTable(query);

            //double utilidadNetaHyP = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    utilidadNetaHyP = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            //double finanNetoHyP = 0.00;

            //#endregion

            //#region 1041_UTILIDAD NETA REFACCIONES

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1041_UTILIDAD NETA REFACCIONES
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1041)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1041_UTILIDAD NETA REFACCIONES
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1041)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}

            //dtInfo = _db.GetDataTable(query);

            //double utilidadNetaRE = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    utilidadNetaRE = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            //#endregion

            //#region 1087_FINANCIAMIENTO NETO REFACCIONES

            //if (idsSuc != "")
            //{
            //    //V1
            //    //1087_FINANCIAMIENTO NETO REFACCIONES
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1087)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}
            //else
            //{
            //    //V1
            //    //1087_FINANCIAMIENTO NETO REFACCIONES
            //    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
            //    "FROM [PREFIX]FINA.FNDRESMEN \r\n" +
            //    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1087)\r\n" +
            //    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            //}

            //dtInfo = _db.GetDataTable(query);

            //double finanNetoRE = 0.00;

            //if (dtInfo.Rows.Count != 0)
            //    finanNetoRE = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            //#endregion

            //double gastosAdminXDep = 0.00;
            //double utilidadNetaXDep = 0.00;

            //List<GADPorcentajesXDepto> lstGADPorc = GADPorcentajesXDepto.Listar(_db, idAgencia);

            //foreach (GADPorcentajesXDepto gadPorc in lstGADPorc)
            //{
            //    if (gadPorc.IdDepartamento == 1) //1_AUTOS NUEVOS
            //    {
            //        #region 1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS

            //        gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

            //        //V1
            //        //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS
            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";

            //        _db.SetQuery(query);

            //        utilidadNetaXDep = utilidadNetaAN + finanNetoAN - gastosAdminXDep;

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";
            //        }

            //        _db.SetQuery(query);


            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";

            //        _db.SetQuery(query);

            //        #endregion
            //    }
            //    else if (gadPorc.IdDepartamento == 51) //51_GERENCIA DE NEGOCIOS
            //    {
            //        #region 1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS

            //        gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";

            //        _db.SetQuery(query);

            //        utilidadNetaXDep = utilidadNetaGN + finanNetoGN - gastosAdminXDep;

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";

            //        _db.SetQuery(query);

            //        #endregion
            //    }
            //    else if (gadPorc.IdDepartamento == 4) //4_AUTOS USADOS
            //    {
            //        #region 1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS

            //        gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";

            //        _db.SetQuery(query);

            //        utilidadNetaXDep = utilidadNetaAU + finanNetoAU - gastosAdminXDep;

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";

            //        _db.SetQuery(query);

            //        #endregion
            //    }
            //    else if (gadPorc.IdDepartamento == 5) //5_SERVICIO
            //    {
            //        #region 1091_GASTOS DE ADMINISTRACION SERVICIO

            //        gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";

            //        _db.SetQuery(query);

            //        utilidadNetaXDep = utilidadNetaSE + finanNetoSE - gastosAdminXDep;

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";

            //        _db.SetQuery(query);

            //        #endregion
            //    }
            //    else if (gadPorc.IdDepartamento == 6) //6_HYP
            //    {
            //        #region 1092_GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA

            //        gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";

            //        _db.SetQuery(query);

            //        utilidadNetaXDep = utilidadNetaHyP + finanNetoHyP - gastosAdminXDep;

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";

            //        _db.SetQuery(query);

            //        #endregion
            //    }
            //    else if (gadPorc.IdDepartamento == 8) //8_REFACCIONES
            //    {
            //        #region 1093_GASTOS DE ADMINISTRACION REFACCIONES

            //        gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";
            //        }
            //        else
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";

            //        _db.SetQuery(query);

            //        utilidadNetaXDep = utilidadNetaRE + finanNetoRE - gastosAdminXDep;

            //        if (idsSuc != "")
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";
            //        }
            //        else                    
            //        {
            //            query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";
            //        }

            //        _db.SetQuery(query);

            //        query = "UPDATE [PREFIX]FINA.FNDRESMEN \r\n" +
            //            "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
            //            "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";

            //        _db.SetQuery(query);

            //        #endregion
            //    }
            //}

            #endregion

            #region V2

            #region 1044_GASTOS ADMINISTRATIVOS

            if (idsSuc != "")
            {
                //V1
                //1044_GASTOS ADMINISTRATIVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1044_GASTOS ADMINISTRATIVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            DataTable dtInfo = _db.GetDataTable(query);

            double gastosAdmin = 0.00;

            if (dtInfo.Rows.Count != 0)
                gastosAdmin = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1008_UTILIDAD NETA DEPARTAMENTAL AUTOS NUEVOS

            if (idsSuc != "")
            {
                //V1
                //1008_UTILIDAD NETA DEPARTAMENTAL AUTOS NUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1008)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1008_UTILIDAD NETA DEPARTAMENTAL AUTOS NUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1008)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaAN = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaAN = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1085_FINANCIAMIENTO NETO AUTOS NUEVOS

            if (idsSuc != "")
            {
                //V1
                //1085_FINANCIAMIENTO NETO AUTOS NUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1085)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1085_FINANCIAMIENTO NETO AUTOS NUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1085)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double finanNetoAN = 0.00;

            if (dtInfo.Rows.Count != 0)
                finanNetoAN = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1012_UTILIDAD NETA GERENCIA DE NEGOCIOS

            if (idsSuc != "")
            {
                //V1
                //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1012)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1012)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaGN = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaGN = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            double finanNetoGN = 0.00;

            #endregion

            #region 1020_UTILIDAD NETA DEPARTAMENTAL

            if (idsSuc != "")
            {
                //V1
                //1020_UTILIDAD NETA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1020)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1020_UTILIDAD NETA DEPARTAMENTAL
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1020)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaAU = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaAU = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS

            if (idsSuc != "")
            {
                //V1
                //1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1086)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1086)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double finanNetoAU = 0.00;

            if (dtInfo.Rows.Count != 0)
                finanNetoAU = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1024_UTILIDAD NETA SERVICIO

            if (idsSuc != "")
            {
                //V1
                //1024_UTILIDAD NETA SERVICIO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1024)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1024_UTILIDAD NETA SERVICIO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1024)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaSE = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaSE = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            double finanNetoSE = 0.00;

            #endregion

            #region 1028_UTILIDAD NETA HOJALATERIA Y PINTURA

            if (idsSuc != "")
            {
                //V1
                //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1028)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1028)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaHyP = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaHyP = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            double finanNetoHyP = 0.00;

            #endregion

            #region 1029_UTILIDAD NETA SERVICIOS ADICIONALES

            //V2
            //1029_UTILIDAD NETA SERVICIOS ADICIONALES
            if (idsSuc != "")
            {
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1029)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1029)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaServAdic = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaServAdic = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion
            
            #region 1041_UTILIDAD NETA REFACCIONES

            if (idsSuc != "")
            {
                //V1
                //1041_UTILIDAD NETA REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1041)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1041_UTILIDAD NETA REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1041)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaRE = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaRE = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1087_FINANCIAMIENTO NETO REFACCIONES

            if (idsSuc != "")
            {
                //V1
                //1087_FINANCIAMIENTO NETO REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1087)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V1
                //1087_FINANCIAMIENTO NETO REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1087)\r\n" +
                "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double finanNetoRE = 0.00;

            if (dtInfo.Rows.Count != 0)
                finanNetoRE = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            double gastosAdminXDep = 0.00;
            double utilidadNetaXDep = 0.00;
            double utilidadNetaServicio = 0.00;

            List<GADPorcentajesXDepto> lstGADPorc = GADPorcentajesXDepto.Listar(_db, idAgencia);

            foreach (GADPorcentajesXDepto gadPorc in lstGADPorc)
            {
                if (gadPorc.IdDepartamento == 1) //1_AUTOS NUEVOS
                {
                    #region 1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS

                    gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

                    //V1
                    //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS
                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1088";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNetaAN + finanNetoAN - gastosAdminXDep;

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";
                    }

                    _db.SetQuery(query);


                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2024";

                    _db.SetQuery(query);

                    #endregion
                }
                else if (gadPorc.IdDepartamento == 51) //51_GERENCIA DE NEGOCIOS
                {
                    #region 1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS

                    gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1089";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNetaGN + finanNetoGN - gastosAdminXDep;

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2025";

                    _db.SetQuery(query);

                    #endregion
                }
                else if (gadPorc.IdDepartamento == 4) //4_AUTOS USADOS
                {
                    #region 1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS

                    gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1090";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNetaAU + finanNetoAU - gastosAdminXDep;

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2026";

                    _db.SetQuery(query);

                    #endregion
                }
                else if (gadPorc.IdDepartamento == 5) //5_SERVICIO
                {
                    #region 1091_GASTOS DE ADMINISTRACION SERVICIO

                    gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1091";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNetaSE + finanNetoSE - gastosAdminXDep;

                    utilidadNetaServicio = utilidadNetaXDep;

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2027";

                    _db.SetQuery(query);

                    #endregion
                }
                else if (gadPorc.IdDepartamento == 6) //6_HYP
                {
                    #region 1092_GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA

                    gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1092";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNetaHyP + finanNetoHyP - gastosAdminXDep;

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2028";

                    _db.SetQuery(query);

                    double utilidadServicioYHyP = utilidadNetaServicio + utilidadNetaXDep + utilidadNetaServAdic;

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1030";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1030";
                    }
                    
                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadServicioYHyP + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1030";

                    _db.SetQuery(query);

                    #endregion
                }
                else if (gadPorc.IdDepartamento == 8) //8_REFACCIONES
                {
                    #region 1093_GASTOS DE ADMINISTRACION REFACCIONES

                    gastosAdminXDep = gastosAdmin * (gadPorc.Porcentaje / 100);

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + gastosAdminXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1093";

                    _db.SetQuery(query);

                    utilidadNetaXDep = utilidadNetaRE + finanNetoRE - gastosAdminXDep;

                    if (idsSuc != "")
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";
                    }
                    else
                    {
                        query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";
                    }

                    _db.SetQuery(query);

                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                        "SET FIFNVALUE = " + utilidadNetaXDep + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 2029";

                    _db.SetQuery(query);

                    #endregion
                }
            }

            #endregion
        }

        public void RecalculaGastosAdministacionYOtrosGastosProductos(int idAgencia, int anio, int mes)
        {
            string idsSuc = "";

            query = "SELECT * FROM [PREFIX]FINA . FNDAGSUC WHERE FIFNIDCIAU = " + idAgencia + " AND FIFNSTATUS = 1 ORDER BY FIFNIDCIAS";

            DataTable dtSucursales = _db.GetDataTable(query);

            if (dtSucursales.Rows.Count != 0)
            {
                foreach (DataRow drSuc in dtSucursales.Rows)
                {
                    idsSuc += drSuc["FIFNIDCIAS"].ToString() + ",";
                }
                idsSuc = idsSuc.Remove(idsSuc.Length - 1, 1);
            }

            PeriodoContable periodo = PeriodoContable.BuscarPorMesAnio(_db, mes, anio);

            SaldoPorPeriodoPorCuentaBalanza saldo6000001000070495 = SaldoPorPeriodoPorCuentaBalanza.Buscar(_db, idAgencia, periodo.Id, 1, 166673); //166673_6000001000070495_GASTOS || ADMINISTRACION || GASTOS GENERALES || INGRESOS ASIMILADOS A SALARIOS DG ||

            double saldoCuenta = 0.00;

            if (saldo6000001000070495 != null)
            {
                saldoCuenta += Convert.ToInt32(saldo6000001000070495.TotalDeCargos - saldo6000001000070495.TotalDeAbonos);

                #region 1044_GASTOS ADMINISTRATIVOS

                if (idsSuc != "")
                {
                    //V1
                    //1044_GASTOS ADMINISTRATIVOS
                    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                        "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
                        "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
                }
                else
                {
                    //V1
                    //1044_GASTOS ADMINISTRATIVOS
                    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                        "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
                        "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
                }

                DataTable dtInfo = _db.GetDataTable(query);

                double gastosAdmin = 0.00;

                if (dtInfo.Rows.Count != 0)
                    gastosAdmin = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                gastosAdmin = gastosAdmin - saldoCuenta;

                if (idsSuc != "")
                {
                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                    "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1044";
                }
                else
                {
                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                    "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1044";
                }

                _db.SetQuery(query);

                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                    "SET FIFNVALUE = " + gastosAdmin + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1044";

                _db.SetQuery(query);

                #endregion

                #region 1049_OTROS (GASTOS) PRODUCTOS

                if (idsSuc != "")
                {
                    //V2
                    //1049_OTROS (GASTOS) PRODUCTOS
                    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                        "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1049)\r\n" +
                        "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
                }
                else
                {
                    //V2
                    //1049_OTROS (GASTOS) PRODUCTOS
                    query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                        "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                        "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1049)\r\n" +
                        "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
                }

                dtInfo = _db.GetDataTable(query);

                double otrosGastosProductos = 0.00;

                if (dtInfo.Rows.Count != 0)
                    otrosGastosProductos = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

                //if (otrosGastosProductos <= 0)
                //    otrosGastosProductos = (otrosGastosProductos * -1) + saldoCuenta;
                //else
                //    otrosGastosProductos = otrosGastosProductos + saldoCuenta;

                otrosGastosProductos = otrosGastosProductos - saldoCuenta;

                if (idsSuc != "")
                {
                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                    "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1049";
                }
                else
                {
                    query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                    "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1049";
                }

                _db.SetQuery(query);

                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                    "SET FIFNVALUE = " + otrosGastosProductos + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1049";

                _db.SetQuery(query);

                #endregion
            }
        }

        public void RecalculaUtilidadDeOperacion(int idAgencia, int anio, int mes)
        {
            string idsSuc = "";

            query = "SELECT * FROM [PREFIX]FINA . FNDAGSUC WHERE FIFNIDCIAU = " + idAgencia + " AND FIFNSTATUS = 1 ORDER BY FIFNIDCIAS";

            DataTable dtSucursales = _db.GetDataTable(query);

            if (dtSucursales.Rows.Count != 0)
            {
                foreach (DataRow drSuc in dtSucursales.Rows)
                {
                    idsSuc += drSuc["FIFNIDCIAS"].ToString() + ",";
                }

                idsSuc = idsSuc.Remove(idsSuc.Length - 1, 1);
            }

            #region 1008_UTILIDAD NETA DEPARTAMENTAL AN

            if (idsSuc != "")
            {
                //V2
                //1008_UTILIDAD NETA DEPARTAMENTAL AN
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1008)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1008_UTILIDAD NETA DEPARTAMENTAL AN
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1008)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            DataTable dtInfo = _db.GetDataTable(query);

            double utilidadNetaAN = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaAN = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1012_UTILIDAD NETA GERENCIA DE NEGOCIOS

            if (idsSuc != "")
            {
                //V2
                //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1012)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1012)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaGN = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaGN = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1020_UTILIDAD NETA DEPARTAMENTAL AU

            if (idsSuc != "")
            {
                //V2
                //1020_UTILIDAD NETA DEPARTAMENTAL AU
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1020)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1020_UTILIDAD NETA DEPARTAMENTAL AU
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1020)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaAU = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaAU = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1024_UTILIDAD NETA SERVICIO

            if (idsSuc != "")
            {
                //V2
                //1024_UTILIDAD NETA SERVICIO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1024)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1024_UTILIDAD NETA SERVICIO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1024)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaServicio = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaServicio = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1028_UTILIDAD NETA HOJALATERIA Y PINTURA

            if (idsSuc != "")
            {
                //V2
                //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1028)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1028)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaHyP = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaHyP = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1029_UTILIDAD NETA SERVICIOS ADICIONALES

            if (idsSuc != "")
            {
                //V2
                //1029_UTILIDAD NETA SERVICIOS ADICIONALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1029)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1029_UTILIDAD NETA SERVICIOS ADICIONALES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1029)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaServAdicionalesServYHyP = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaServAdicionalesServYHyP = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1041_UTILIDAD NETA REFACCIONES

            if (idsSuc != "")
            {
                //V2
                //1041_UTILIDAD NETA REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1041)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1041_UTILIDAD NETA REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1041)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaRE = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaRE = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1045_OTROS INGRESOS PLANTA

            if (idsSuc != "")
            {
                //V2
                //1045_OTROS INGRESOS PLANTA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1045)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1045_OTROS INGRESOS PLANTA
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1045)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double otrosIngresosPlanta = 0.00;

            if (dtInfo.Rows.Count != 0)
                otrosIngresosPlanta = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1044_GASTOS ADMINISTRATIVOS

            if (idsSuc != "")
            {
                //V2
                //1044_GASTOS ADMINISTRATIVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1044_GASTOS ADMINISTRATIVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1044)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double gastosAdministacion = 0.00;

            if (dtInfo.Rows.Count != 0)
                gastosAdministacion = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1046_UTILIDAD DE OPERACION

            double utilidadNetaOperacion = utilidadNetaAN + utilidadNetaGN + utilidadNetaAU + utilidadNetaServicio + utilidadNetaHyP + utilidadNetaServAdicionalesServYHyP 
                + utilidadNetaRE + otrosIngresosPlanta - gastosAdministacion;

            //V2
            //1051_UTILIDAD NETA ANTES DE FIDEICOMISO
            if (idsSuc != "")
            {
                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1046";
            }
            else
            {
                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1046";
            }

            _db.SetQuery(query);

            query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = " + utilidadNetaOperacion + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1046";

            _db.SetQuery(query);

            #endregion
        }

        public void RecalculaUNAFyUFAIyUN(int idAgencia, int anio, int mes)
        {
            string idsSuc = "";

            query = "SELECT * FROM [PREFIX]FINA . FNDAGSUC WHERE FIFNIDCIAU = " + idAgencia + " AND FIFNSTATUS = 1 ORDER BY FIFNIDCIAS";

            DataTable dtSucursales = _db.GetDataTable(query);

            if (dtSucursales.Rows.Count != 0)
            {
                foreach (DataRow drSuc in dtSucursales.Rows)
                {
                    idsSuc += drSuc["FIFNIDCIAS"].ToString() + ",";
                }

                idsSuc = idsSuc.Remove(idsSuc.Length - 1, 1);
            }

            #region 1048_FINANCIAMIENTO NETO

            if (idsSuc != "")
            {
                //V2
                //1048_FINANCIAMIENTO NETO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1048)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1048_FINANCIAMIENTO NETO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1048)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            DataTable dtInfo = _db.GetDataTable(query);

            double finanNeto = 0.00;

            if (dtInfo.Rows.Count != 0)
                finanNeto = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1049_OTROS (GASTOS) PRODUCTOS

            if (idsSuc != "")
            {
                //V2
                //1049_OTROS (GASTOS) PRODUCTOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1049)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1049_OTROS (GASTOS) PRODUCTOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1049)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double otrosGastosProductos = 0.00;

            if (dtInfo.Rows.Count != 0)
                otrosGastosProductos = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1050_GASTOS CORPORATIVOS

            if (idsSuc != "")
            {
                //V2
                //1050_GASTOS CORPORATIVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1050)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1050_GASTOS CORPORATIVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1050)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double gastoCorpo = 0.00;

            if (dtInfo.Rows.Count != 0)
                gastoCorpo = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1046_UTILIDAD DE OPERACION

            if (idsSuc != "")
            {
                //V2
                //1046_UTILIDAD DE OPERACION
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1046)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1046_UTILIDAD DE OPERACION
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1046)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadDeOperacion = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadDeOperacion = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1085_FINANCIAMIENTO NETO AUTOS NUEVOS

            if (idsSuc != "")
            {
                //V2
                //1085_FINANCIAMIENTO NETO AUTOS NUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1085)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1085_FINANCIAMIENTO NETO AUTOS NUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1085)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double finanNetoAN = 0.00;

            if (dtInfo.Rows.Count != 0)
                finanNetoAN = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS

            if (idsSuc != "")
            {
                //V2
                //1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1086)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1086)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double finanNetoAU = 0.00;

            if (dtInfo.Rows.Count != 0)
                finanNetoAU = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1087_FINANCIAMIENTO NETO REFACCIONES

            if (idsSuc != "")
            {
                //V2
                //1087_FINANCIAMIENTO NETO REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1087)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1087_FINANCIAMIENTO NETO REFACCIONES
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1087)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double finanNetoRE = 0.00;

            if (dtInfo.Rows.Count != 0)
                finanNetoRE = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1051_UTILIDAD NETA ANTES DE FIDEICOMISO

            double unaf = utilidadDeOperacion + finanNetoAN + finanNetoAU + finanNetoRE + finanNeto + otrosGastosProductos - gastoCorpo;

            //V2
            //1051_UTILIDAD NETA ANTES DE FIDEICOMISO
            if (idsSuc != "")
            {
                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1051";
            }
            else
            {
                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1051";
            }

            _db.SetQuery(query);

            query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = " + unaf + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1051";

            _db.SetQuery(query);

            #endregion

            #region 1052_UTILIDAD NETA DEL FIDEICOMISO

            if (idsSuc != "")
            {
                //V2
                //1052_UTILIDAD NETA DEL FIDEICOMISO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1052)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1052_UTILIDAD NETA DEL FIDEICOMISO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1052)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaFideicomiso = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaFideicomiso = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1053_UTILIDAD NETA TAXIS BAM

            if (idsSuc != "")
            {
                //V2
                //1053_UTILIDAD NETA TAXIS BAM
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1053)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1053_UTILIDAD NETA TAXIS BAM
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1053)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double utilidadNetaTaxiBAM = 0.00;

            if (dtInfo.Rows.Count != 0)
                utilidadNetaTaxiBAM = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1054_PARTIDAS EXTRAORDINARIAS

            if (idsSuc != "")
            {
                //V2
                //1054_PARTIDAS EXTRAORDINARIAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1054)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1054_PARTIDAS EXTRAORDINARIAS
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1054)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double partidasExtraordinarias = 0.00;

            if (dtInfo.Rows.Count != 0)
                partidasExtraordinarias = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1055_UTILIDAD FINAL ANTES DE IMPUESTOS

            double ufai = unaf + utilidadNetaFideicomiso + utilidadNetaTaxiBAM - partidasExtraordinarias;

            //V2
            //1055_UTILIDAD FINAL ANTES DE IMPUESTOS
            if (idsSuc != "")
            {
                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1055";
            }
            else
            {
                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1055";
            }

            _db.SetQuery(query);

            query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = " + ufai + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1055";

            _db.SetQuery(query);

            #endregion

            #region 1056_I.S.R. CORRIENTE

            if (idsSuc != "")
            {
                //V2
                //1056_I.S.R. CORRIENTE
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1056)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1056_I.S.R. CORRIENTE
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1056)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double isrCorriente = 0.00;

            if (dtInfo.Rows.Count != 0)
                isrCorriente = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1057_I.S.R. DIFERIDO

            if (idsSuc != "")
            {
                //V2
                //1057_I.S.R. DIFERIDO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1057)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1057_I.S.R. DIFERIDO
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1057)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double isrDiferido = 0.00;

            if (dtInfo.Rows.Count != 0)
                isrDiferido = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1058_P.T.U.

            if (idsSuc != "")
            {
                //V2
                //1058_P.T.U.
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1058)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }
            else
            {
                //V2
                //1058_P.T.U.
                query = "SELECT FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD, SUM(FIFNVALUE) VALOR \r\n" +
                    "FROM [PREFIX]FINA.FNDRSMENE \r\n" +
                    "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNSTATUS = 1 AND FIFNCPTD IN (1058)\r\n" +
                    "GROUP BY FIFNYEAR, FIFNMONTH, FIFNCPTD, FSFNCPTD";
            }

            dtInfo = _db.GetDataTable(query);

            double ptu = 0.00;

            if (dtInfo.Rows.Count != 0)
                ptu = Convert.ToDouble(dtInfo.Rows[0]["VALOR"]);

            #endregion

            #region 1059_UTILIDAD NETA

            double utilidadNeta = ufai - isrCorriente - isrDiferido - ptu;

            //V2
            //1059_UTILIDAD NETA
            if (idsSuc != "")
            {
                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + "," + idsSuc + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1059";
            }
            else
            {
                query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = 0, USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1059";
            }

            _db.SetQuery(query);

            query = "UPDATE [PREFIX]FINA.FNDRSMENE \r\n" +
                "SET FIFNVALUE = " + utilidadNeta + ", USERUPDAT='3665', PROGUPDAT='1241', DATEUPDAT=CURRENT_DATE, TIMEUPDAT=CURRENT_TIME \r\n" +
                "WHERE FIFNIDCIAU IN (" + idAgencia + ") AND FIFNYEAR = " + anio + " AND FIFNMONTH = " + mes + " AND FIFNCPTD = 1059";

            _db.SetQuery(query);

            #endregion
        }
    }
}
