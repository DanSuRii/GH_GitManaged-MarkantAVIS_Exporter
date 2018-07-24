using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MarkantAVIS_Exporter
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void btnSelectFiles(object sender, RoutedEventArgs e)
		{
			lbFileSelected.Items.Clear();

			Microsoft.Win32.OpenFileDialog oFD = new Microsoft.Win32.OpenFileDialog();
			oFD.Multiselect = true;
			oFD.InitialDirectory = @"C:\Users\A.Roennburg\Documents\GH_ArbeitPlatz\Arbeits vom Sandra\MARKANT_CalculateSheet\";
			//openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			if(true == oFD.ShowDialog())
			{
				foreach( string fileName in oFD.FileNames )
				{
					//lbFileSelected.Items.Add( System.IO.Path.GetFileName(fileName) );
					lbFileSelected.Items.Add(System.IO.Path.GetFullPath(fileName));
				}
			}
		}

		int RunItem( string strFilePath )
		{
			//System.Diagnostics.ProcessStartInfo pSI = ;
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.StartInfo = new System.Diagnostics.ProcessStartInfo("ExcelWorker.exe", "\""+ strFilePath + "\"");
			proc.Start();

			proc.WaitForExit();
			return proc.ExitCode;
		}
			
		private void btnRun(object sender, RoutedEventArgs e)
		{
			object objLock = new object();
			Queue<string> queueFiles = new Queue<string>();

			foreach ( string cur in lbFileSelected.Items )
			{
				//System.Diagnostics.Debug.WriteLine(cur);
				queueFiles.Enqueue(cur);
			}

			if(true == cbMultiThread.IsChecked)
			{
				List<Task<int>> listTask = new List<Task<int>>();
				System.Threading.ThreadPool.SetMaxThreads(4, 4);

				foreach (string cur in queueFiles)
				{
					listTask.Add(
						Task<int>.Run(
							() => {
								return RunItem(cur);
							}
							)
						);
				}

				System.Threading.Tasks.Task.WaitAll(listTask.ToArray());
			}
			else
			{
				foreach (string cur in queueFiles)
					RunItem(cur);
			}

		}
	}
}
