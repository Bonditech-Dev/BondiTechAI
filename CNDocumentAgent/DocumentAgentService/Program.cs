using System.Threading;

namespace DocumentAgentService
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        private static void Main()
        {
#if DEBUG
            DocumentAgentService docAgent = new DocumentAgentService();
            docAgent.DebugOnStart();
            Thread.Sleep(Timeout.Infinite);
#else
                  ServiceBase[] ServicesToRun;
                  ServicesToRun = new ServiceBase[]
                  {
                      new DocumentAgentService()
                  };
                  ServiceBase.Run(ServicesToRun);
#endif
        }
    }
}