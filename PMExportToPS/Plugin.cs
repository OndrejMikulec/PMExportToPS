/*
 * Created by SharpDevelop.
 * User: Ondra
 * Date: 10/04/2016
 * Time: 07:08
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 * 
 * 
 * Instalation:
 * 		In administrators cmd go to the PMExportToPS.dll directory.
 *		C:\WINDOWS\Microsoft.NET\Framework64\v4.0.30319\regasm.exe PMExportToPS.dll /register /codebase
 *
 *		REG ADD "HKCR\CLSID\{db6fda6f-c255-423c-93b1-a80001c9ce25}\Implemented Categories\{311b0135-1826-4a8c-98de-f313289f815e}" /reg:64 /f
 * 
 *Unistalation:
* 
* 		REG DELETE "HKCR\CLSID\{db6fda6f-c255-423c-93b1-a80001c9ce25}\Implemented Categories\{311b0135-1826-4a8c-98de-f313289f815e}" /reg:64 /f
* 
* 
* 
* 
* 
 * Plugin run:
 * 		Plugin {db6fda6f-c255-423c-93b1-a80001c9ce25} run
 * 
 * * Surrogate connect:
 * 		Plugin {db6fda6f-c255-423c-93b1-a80001c9ce25} debug
 * 		Plugin {db6fda6f-c255-423c-93b1-a80001c9ce25} debugLoad
 * 
 */
using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;
using System.Xml;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace PMExportToPS
{
	[Guid("db6fda6f-c255-423c-93b1-a80001c9ce25")] 
	[ClassInterface(ClassInterfaceType.None)] 
	[ComVisible(true)]
	public class Plugin : PowerMILL.IPowerMILLPlugin
	{
		string m_token; 
		PowerMILL.PluginServices m_services;  
		int m_parent_window;
		Options m_form = null;
		
		public string ThisPath
		{
			get {
				return System.Reflection.Assembly.GetExecutingAssembly().Location;
			}
		}
		
		public string SerializePath
		{
			get {
				return Path.GetDirectoryName(ThisPath) + @"\PMExportToPS.xml";;
			}
		}
		
		
		
		string _pathPS;
		public string PathPS {
			get {
				return _pathPS;
			}
			set {
				_pathPS = value;
			}
		}
		
		public Plugin()
		{
		}

		#region IPowerMILLPlugin implementation

		public void PreInitialise(string locale)
		{
			
		}

		public void Initialise(string Token, PowerMILL.PluginServices pServices, int ParentWindow)
		{
			m_token =Token;
			m_services = pServices;
			m_parent_window = ParentWindow;
			
			
			MySerialization.Load(this);
			
		}

		public void Uninitialise()
		{
			MySerialization.Save(this);
			
			 m_services = null;    
			 GC.Collect();
		}

		public void Version(out int pMajor, out int pMinor, out int pIssue)
		{
			pMajor = 0;    
			 pMinor = 0;    
			 pIssue = 0;
		}

		public void MinPowerMILLVersion(out int pMajor, out int pMinor, out int pIssue)
		{
			pMajor = 13;   
			 pMinor = 1;    
			 pIssue = 0;
		}

		public void DisplayOptions()
		{
			RaiseForm();
		}

		public void PluginIconBitmap(out PowerMILL.PluginBitmapFormat pFormat, out byte[] pPixelData, out int pWidth, out int pHeight)
		{
			
			throw new NotImplementedException();
		}

		public void SerializeProjectData(string Path, bool Saving)
		{
			throw new NotImplementedException();
		}

		public void ProcessEvent(string EventData)
		{
            throw new NotImplementedException();

		}

		public void ProcessCommand(string Command)
		{
			if (Command.ToLower().Trim()=="run") {
				DoWork();
			}

            if (Command.ToLower().Trim() == "debug")
            {
                System.Diagnostics.Debugger.Break();
            }

            if (Command.ToLower().Trim() == "debugload")
            {
                System.Diagnostics.Debugger.Break();
                MySerialization.Load(this);
            }
        }
		
		async void DoWork()
		{
			while (m_services.Busy) {
				await Task.Delay(100);
			}
			
			using (PMInteraction inter = new PMInteraction(m_token,m_services)) {
				new DoWork(m_token,m_services,_pathPS);
			}
		}
		

		public string Name {
			get {
				return "Export to PowerShape";
			}
		}

		public string Description {
			get {
				throw new NotImplementedException();
			}
		}

		public string Author {
			get {
				return "Ondrej Mikulec";
			}
		}

		public bool HasOptions {
			get {
				return true;
			}
		}

		#endregion
		
		void RaiseForm() { 
			
			m_form = new Options(m_token,m_services,this);   

		   WinFormWrapper _oWrp = new WinFormWrapper(m_parent_window);
		   m_form.Show(_oWrp);
		   
		}
		

	}

	
}
