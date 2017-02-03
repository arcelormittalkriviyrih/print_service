using System;
using System.ServiceProcess;

namespace PrintWindowsService
{
	public partial class PrintService : ServiceBase
	{
		private PrintJobs pJobs;

        #region Constructor

        public PrintService()
		{
			InitializeComponent();
            RequestAdditionalTime(60000);//60 sec
            // Set up a timer to trigger every print task frequency.
            pJobs = new PrintJobs();
        }

        #endregion

        #region Methods

        protected override void OnStart(string[] args)
		{
            pJobs.StartJob();
        }

		protected override void OnStop()
		{
            pJobs.StopJob();
        }
        #endregion
    }
}
