using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Threading;
using System.Net;
using System.Net.Sockets;

/************************************************************************/
/* Autor: J. Antonio Barrera F.											*/
/************************************************************************/
namespace DemonioPAB
{
    public class ServidorAsincrono
    {
        //public DVAControls.csLog m_log = null;

        private long SegundosEspera = 0;
        private int puerto = 0;
        private TcpListener listener;
        private Thread listenerThread;
        DVADB.DB2 dbCnx;

        //public Boolean escribeLOG;

        public ServidorAsincrono(int PuertoEscuchar, long TiempoAbortar, ref DVADB.DB2 _dbCnx)
        {
            this.dbCnx = _dbCnx;
            puerto = PuertoEscuchar;
            SegundosEspera = TiempoAbortar;
            if (Ejecutor.escribeLOG)
            {
                Ejecutor.m_log.AgregaRegistro("ServidorAsincrono.ServidorAsincrono()");
                Ejecutor.m_log.AgregaRegistro("Puerto: " + puerto);
                Ejecutor.m_log.AgregaRegistro("SegundosEspera: " + SegundosEspera);
            }
        }

        public void Iniciar()
        {
            if (Ejecutor.escribeLOG)
                Ejecutor.m_log.AgregaRegistro("ServidorAsincrono.Iniciar()");

            listenerThread = new Thread(new ThreadStart(DoListen));
            listenerThread.Start();
        }

        private void DoListen()
        {
            if (Ejecutor.escribeLOG)
                Ejecutor.m_log.AgregaRegistro("ServidorAsincrono.DoListen()");

            //try
            //{
            //JASG
            //IPAddress direc = IPAddress.Parse("10.1.4.92");            
            //listener = new TcpListener(direc, puerto);
            listener = new TcpListener(System.Net.IPAddress.Any, puerto);
            listener.Start();                

                do
                {
                    UserConnection client = new UserConnection(listener.AcceptTcpClient(), 2000, ref dbCnx);
                    //client.m_log = m_log;
                    client.LineReceived += new LineReceive(OnLineReceived);
                } 
                while (true);
            //}
            //catch(Exception ex)
            //{
            //    m_log.AgregaRegistro("Message: " + ex.Message);
            //    m_log.AgregaRegistro("InnerException: " + ex.InnerException);
            //    m_log.AgregaRegistro("StackTrace: " + ex.StackTrace);

            //    Console.WriteLine(ex.Message);
            //}
        }

        private void OnLineReceived(UserConnection sender, String data)
        {
            if (Ejecutor.escribeLOG)
                Ejecutor.m_log.AgregaRegistro("ServidorAsincrono.OnLineReceived(UserConnection sender, String data)");

            System.Windows.Forms.MessageBox.Show(data);
            Console.WriteLine(data);
        }

        public Boolean Detener()
        {
            if (Ejecutor.escribeLOG)
                Ejecutor.m_log.AgregaRegistro("ServidorAsincrono.Detener()");

            try
            {
                DateTime inicia = DateTime.Now;
                listenerThread.Abort();

                return true;
            }
            catch (Exception ex)
            {
                if (Ejecutor.escribeLOG)
                {
                    Ejecutor.m_log.AgregaRegistro("Message: " + ex.Message);
                    Ejecutor.m_log.AgregaRegistro("InnerException: " + ex.InnerException);
                    Ejecutor.m_log.AgregaRegistro("StackTrace: " + ex.StackTrace);
                }
                                
                return false;
            }
        }
    }
}
