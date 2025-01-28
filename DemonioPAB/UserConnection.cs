using System;
using System.Net.Sockets;
using System.Text;
using System.IO;
using System.Threading;

namespace DemonioPAB
{
    /************************************************************************/
    /* Autor: J. Antonio Barrera F.											*/
    /************************************************************************/
    public delegate void LineReceive(UserConnection sender, string Data);

    /**
    * Funciona como un TcpClient connection para un usuario simple
    **/
    public class UserConnection
    {
        public DVAControls.csLog m_log = null;

        private const int READ_BUFFER_SIZE = 2048;//2Kb
        private TcpClient client;
        private byte[] readBuffer = new byte[READ_BUFFER_SIZE];
        public event LineReceive LineReceived;
        DVADB.DB2 dbConex;

        public UserConnection(TcpClient client, int bufferLectura, ref DVADB.DB2 dbConex)
        {
            this.dbConex = dbConex;
            if (Ejecutor.escribeLOG)
                Ejecutor.m_log.AgregaRegistro("UserConnection.UserConnection(TcpClient client, int bufferLectura)" + client.Client.RemoteEndPoint.ToString());
            
            this.client = client;
            this.client.GetStream().BeginRead(readBuffer, 0, READ_BUFFER_SIZE, new AsyncCallback(StreamReceiver), null);
        }

        /**
         * Inicia la lectura asincrona del Stream
         **/
        private void StreamReceiver(IAsyncResult ar)
        {
            if (Ejecutor.escribeLOG)
                Ejecutor.m_log.AgregaRegistro("UserConnection.StreamReceiver(IAsyncResult ar)");

            int BytesRead;

            //try
            //{
                BytesRead = client.GetStream().EndRead(ar);

                SolicitudServicio SS = new SolicitudServicio(readBuffer, client , ref dbConex);
  /*              Thread AtiendePeticion = new Thread(SS.atiendeSolicitud);
                AtiendePeticion.Start();

                
                //me aseguro de que no haya otro hilo usando el stream al mismo tiempo
                lock (client.GetStream())
                {
                    //Inicia una lectura asincrona del buffer
                    client.GetStream().BeginRead(readBuffer, 0, READ_BUFFER_SIZE, new AsyncCallback(StreamReceiver), null);
                }*/
            //}
            //catch (Exception ex)
            //{
            //    m_log.AgregaRegistro("Message: " + ex.Message);
            //    m_log.AgregaRegistro("InnerException: " + ex.InnerException);
            //    m_log.AgregaRegistro("StackTrace: " + ex.StackTrace);

            //    Console.WriteLine(ex.Message);
            //}
        }
    }
}