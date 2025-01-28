namespace WorkerServiceDPAB
{
    public class Worker : BackgroundService
    {
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                // Aqu� coloca la l�gica principal de tu EXE
                Log("Servicio ejecut�ndose... " + DateTime.Now);
                await Task.Delay(10000, stoppingToken); // Intervalo entre tareas
            }
        }

        private void Log(string message)
        {
            string path = "C:\\Logs\\MyService.log";
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.AppendAllText(path, $"{DateTime.Now}: {message}{Environment.NewLine}");
        }
    }
}
