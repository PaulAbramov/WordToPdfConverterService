using System.ServiceProcess;

namespace WordToPdfConverter
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] _arguments)
        {
            var ServicesToRun = new ServiceBase[]
            {
                new DocxToPdfConverterService(_arguments)
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
