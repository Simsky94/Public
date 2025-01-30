using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Threading;
using System.Net;
using System.Net.Sockets;

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
        private CancellationTokenSource cts = new CancellationTokenSource();
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

            if (listenerThread != null && listenerThread.IsAlive)
            {
                Ejecutor.m_log.AgregaRegistro("El servidor ya está en ejecución.");
                return; // Evita iniciar múltiples veces
            }

            listenerThread = new Thread(new ThreadStart(DoListen));
            listenerThread.Start();
        }

        private void DoListen()
        {
            if (Ejecutor.escribeLOG)
                Ejecutor.m_log.AgregaRegistro("ServidorAsincrono.DoListen()");

            listener = new TcpListener(System.Net.IPAddress.Any, puerto);

            // Habilita SO_REUSEADDR para evitar el error de puerto en uso
            listener.Server.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);

            listener.Start();

            try
            {
                while (!cts.Token.IsCancellationRequested)
                {
                    if (listener.Pending()) // Evita bloqueos en AcceptTcpClient
                    {
                        UserConnection client = new UserConnection(listener.AcceptTcpClient(), 2000, ref dbCnx);
                        client.LineReceived += new LineReceive(OnLineReceived);
                    }
                    Thread.Sleep(100); // Pequeña espera para reducir CPU
                }
            }
            catch (SocketException ex)
            {
                Ejecutor.m_log.AgregaRegistro("Error en el listener: " + ex.Message);
            }

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
                cts.Cancel(); // Indica que se debe detener el listener

                if (listener != null)
                {
                    listener.Stop();
                    listener = null;
                }

                return true;
            }
            catch (Exception ex)
            {
                if (Ejecutor.escribeLOG)
                {
                    Ejecutor.m_log.AgregaRegistro("Error al detener el servidor: " + ex.Message);
                }
                return false;
            }
        }

    }
}
