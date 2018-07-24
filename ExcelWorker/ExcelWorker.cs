using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using RMarschal = System.Runtime.InteropServices.Marshal;

namespace ExcelWorker
{
	using Callable = Func<bool>;

	public interface IRAIIBase
	{
		void DoIt();
	}

	public class RAII<T> : IRAIIBase
	{
		private Func<T, bool> fnExit;
		private T obj;

		public RAII(Func<T, bool> fnExit, T obj)
		{
			this.fnExit = fnExit;
			this.obj = obj;
		}

		public void DoIt()
		{
			fnExit(obj);
		}


	}

/*
	class RAIICont : IDisposable
	{
		//public delegate void Callable();


		public void Dispose()
		{
			foreach (IRAIIBase cur in queueRaii)
			{
				cur.DoIt();
			}
		}

		public void Push(IRAIIBase f)
		{
			queueRaii.Enqueue(f);
		}

		private Queue<IRAIIBase> queueRaii = new Queue<IRAIIBase>();
	}
 */

	/*
	 */
		class RAIICont : IDisposable
		{
			//public delegate void Callable();


			public void Dispose()
			{
				foreach( Callable cur in stackRaii ) {
					cur();
				}
			}

			public void Push( Callable f )
			{
				stackRaii.Push(f);
			}

			//private Queue< Func<bool> > queueRaii; <<-- Falsch here
			private Stack< Func<bool> > stackRaii = new Stack< Func<bool> >();
		}

	class RAIIFuncs
	{
		public Func<bool> GetFunc( Excel.Application appXls )
		{
			return () => {
				appXls.Quit();
				RMarschal.ReleaseComObject(appXls);
				return true;
			};
		}

		public Func<bool> GetFunc(Excel.Workbook oWB)
		{
			return () => {
				oWB.Close(SaveChanges:false);
				RMarschal.ReleaseComObject(oWB);
				return true;
			};
		}


		public Func<bool> GetFunc<T>(T marsObj)
		{
			return () => {
				RMarschal.ReleaseComObject(marsObj);
				return true;
			};
		}

	}

	class ExcelWorker
	{
		public void SetPrintAreaToTable( Excel.Worksheet oWS, string strTblName )
		{
			RAIIFuncs rFuncs = new RAIIFuncs();
			using (RAIICont _raiiCont = new RAIICont())
			{
				Excel.ListObjects oLOs = oWS.ListObjects;
				_raiiCont.Push(rFuncs.GetFunc(oLOs));

				Excel.ListObject oLO = oLOs[strTblName];
				_raiiCont.Push(rFuncs.GetFunc(oLO));

				Excel.Range rngLO = oLO.Range;
				_raiiCont.Push(rFuncs.GetFunc(rngLO));

				string prntArea = rngLO.Address;

				Excel.PageSetup pageSetup = oWS.PageSetup;
				//_raiiCont.Push(rFuncs.GetFunc(pageSetup));

				pageSetup.PrintArea = prntArea;
			}
		}

		public void SetPrnAreaToSingleWidth( Excel.Application oXL, Excel.Worksheet oWS )
		{
			RAIIFuncs rFuncs = new RAIIFuncs();
			using (RAIICont _raiiCont = new RAIICont())
			{
				//oXL.PrintCommunication = false;

				Excel.PageSetup ps = oWS.PageSetup;
				//_raiiCont.Push(rFuncs.GetFunc(ps));

				{//header and footer setup
					ps.LeftHeader = "";
					ps.CenterHeader = "";
					ps.RightHeader = "";
					ps.LeftFooter = "";
					ps.CenterFooter = "";
					ps.RightFooter = "";
				}//end of header and footer setup

				{//begin of margin
					ps.LeftMargin = oXL.InchesToPoints(0.7);
					ps.RightMargin = oXL.InchesToPoints(0.7);
					ps.TopMargin = oXL.InchesToPoints(0.78);
					ps.BottomMargin = oXL.InchesToPoints(0.78);
					ps.HeaderMargin = oXL.InchesToPoints(0.3);
					ps.FooterMargin = oXL.InchesToPoints(0.3);
				}//end of margin

				{//etcs
					ps.PrintHeadings = false;
					ps.PrintGridlines = false;
					ps.PrintComments = Excel.XlPrintLocation.xlPrintNoComments;
					ps.CenterHorizontally = false;
					ps.CenterVertically = false;
					ps.Orientation = Excel.XlPageOrientation.xlLandscape;
					ps.Draft = false;
					ps.PaperSize = Excel.XlPaperSize.xlPaperA4;
					ps.FirstPageNumber = (int)Excel.Constants.xlAutomatic;
					ps.Order = Excel.XlOrder.xlDownThenOver;
					ps.BlackAndWhite = false;
					ps.Zoom = false;
					ps.PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed;
					ps.OddAndEvenPagesHeaderFooter = false;
					ps.DifferentFirstPageHeaderFooter = false;
					ps.ScaleWithDocHeaderFooter = true;
					ps.AlignMarginsHeaderFooter = true;					
				}//etcs


				ps.FitToPagesWide = 1;
				ps.FitToPagesTall = false;


				oXL.PrintCommunication = true;
			}
		}

		public void ExportCSV( Excel.Application oAppXls ,Excel.Workbooks oWBBs ,Excel.Worksheet oWS, string strTableName, string strFileName )
		{
			RAIIFuncs rFuncs = new RAIIFuncs();
			using (RAIICont _raiiCont = new RAIICont())
			{
				Excel.Workbook _destWB = oWBBs.Add();
				_raiiCont.Push(rFuncs.GetFunc(_destWB));

				Excel.Sheets _oWSs = _destWB.Sheets;
				_raiiCont.Push(rFuncs.GetFunc(_oWSs));

				Excel.Worksheet _destWS = _oWSs[1];
				_raiiCont.Push(rFuncs.GetFunc(_destWS));

				Excel.Range _rngDest = _destWS.Range["A1"];
				_raiiCont.Push(rFuncs.GetFunc(_rngDest));

				Excel.ListObjects _LOs = oWS.ListObjects;
				_raiiCont.Push(rFuncs.GetFunc(_LOs));

				Excel.ListObject _srcLO = _LOs[strTableName];
				_raiiCont.Push(rFuncs.GetFunc(_srcLO));

				Excel.Range _srcRng = _srcLO.Range;
				_raiiCont.Push(rFuncs.GetFunc(_srcRng));

				_srcRng.Copy(_rngDest);

				oAppXls.DisplayAlerts = false;

				_destWB.SaveAs(
					Filename: @"C:\Users\A.Roennburg\Documents\GH_ArbeitPlatz\Arbeits vom Sandra\MARKANT_CalculateSheet\CSV Daten\" + strFileName + ".csv"
					, FileFormat: Excel.XlFileFormat.xlCSV
					, CreateBackup: false
					, Local: true
					);

			}
				oAppXls.DisplayAlerts = true;
		}

		/*
			Do not Parrelism here.
			It will be controlled by ... Program, it just callee
		*/
		public int EntryPoint( string pathXls )
		{
			RAIIFuncs rFuncs = new RAIIFuncs();
			using ( RAIICont _raiiCont = new RAIICont() )
			{
				Excel.Application appExcel = new Excel.Application();

				#region Removed_Code
				/*
		Func<bool> toCall = () =>
		{
			appExcel.Quit();
			RMarschal.ReleaseComObject(appExcel);
			return true;
		};
		 */
				/*
				_raiiCont.Push( new RAII<Excel.Application>( 
					x => {
						x.Quit();
						RMarschal.ReleaseComObject(x);
						return true;
					}
					, appExcel
					) );
				 */

				//throw new System.Exception("OMG");
				#endregion

				_raiiCont.Push( rFuncs.GetFunc(appExcel) );

				Excel.Workbooks oWBs = appExcel.Workbooks;
				_raiiCont.Push(rFuncs.GetFunc(oWBs));


				Excel.Workbook _WB = oWBs.Open(pathXls, ReadOnly:true);
				_raiiCont.Push(rFuncs.GetFunc(_WB));

				string strName = _WB.Name;
				string strBaseName = System.IO.Path.GetFileNameWithoutExtension(_WB.FullName);
				//System.Diagnostics.Debug.WriteLine( strBaseName );

				Excel.Sheets oWSs = _WB.Sheets;
				_raiiCont.Push(rFuncs.GetFunc(oWSs));

				Excel.Worksheet _WS = oWSs["ARBETISTABELLE"];
				_raiiCont.Push(rFuncs.GetFunc(_WS));

				SetPrintAreaToTable(_WS, "Tabelle1");
				SetPrnAreaToSingleWidth(appExcel, _WS);

				_WS.ExportAsFixedFormat(
					Type: Excel.XlFixedFormatType.xlTypePDF
					, Filename: @"C:\Users\A.Roennburg\Documents\GH_ArbeitPlatz\Arbeits vom Sandra\MARKANT_CalculateSheet\PDF Daten\" + strBaseName + ".pdf"
					, Quality: Excel.XlFixedFormatQuality.xlQualityStandard
					, IncludeDocProperties: true
					, IgnorePrintAreas: false
					, OpenAfterPublish: false
					);

				ExportCSV( appExcel,oWBs ,_WS, "Tabelle1", strBaseName);

			}

			return 0;
		}
	}
}
