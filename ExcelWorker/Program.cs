using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWorker
{
	class Program
	{
		/// <summary>
		/// Argument only 1 Excel,
		/// List manage by GUI this jsut work Excel Exporter.
		/// Then 4 Trhead work with Task, 
		/// Proc(){  using (syncfileQueue) 
		///			{
		///				FileNameQueue.Pop()
		///			}
		///			Process proc(".../???.xls");
		///			WaitforExit()
		///	}
		/// </summary>
		/// <param name="args"></param>


		static int Main(string[] args)
		{
			if( 0 == args.Length )
			{
				Console.WriteLine("Argument Does not exists");
				return -1;
			}

			string fullPath = "";
			try
			{
				fullPath = System.IO.Path.GetFullPath(args[0]);
			}
			catch(Exception)
			{
				Console.WriteLine("Path does not exists");
				return -2;
			}

			ExcelWorker worker = new ExcelWorker();
			worker.EntryPoint(fullPath);


			return 0;
		}
	}
}
