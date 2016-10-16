/*
 * Created by SharpDevelop.
 * User: Ondra
 * Date: 25/03/2016
 * Time: 16:15
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;



namespace PMExportToPS
{
	/// <summary>
	/// Description of PMInteraction.
	/// </summary>
	public class PMInteraction : IDisposable
	{
		bool sucess = true;
		public bool Sucess {
			get {
				return sucess;
			}
		}
		

		bool _dialogy;
		bool _disableAll = true;
		
				
		string m_token; 
		PowerMILL.PluginServices m_services;


		public PMInteraction(string m_token, PowerMILL.PluginServices m_services) 
		{
			this.m_token = m_token;
			this.m_services = m_services;
			
						
			if (m_services==null) {
				sucess=false;
				return;
			}
			
			Console.WriteLine("SpinWait 5s timeout.........");
			System.Threading.SpinWait.SpinUntil(()=>!m_services.Busy,5000);
			Console.WriteLine("SpinWait end");
			if (m_services.Busy) {
				sucess=false;
				return;
			}

			m_services.DoCommand(m_token,@"QUIT");
			m_services.DoCommand(m_token,@"QUIT");
			m_services.DoCommand(m_token,@"QUIT");
			m_services.DoCommand(m_token,@"ECHO OFF DCPDEBUG UNTRACE COMMAND ACCEPT");
			
			_dialogy = getActualDialogs();
			m_services.DoCommand(m_token,"DIALOGS MESSAGE OFF");
			m_services.DoCommand(m_token,"DIALOGS ERROR OFF");
			
			
			
		}
		
		public bool getActualDialogs()
		{
 			/*int err = 0;
 			string vysledek = null;
 			oAppliacation.ExecuteEx(@"print $Status.Dialog.Message",out err,out vysledek);
 			vysledek = vysledek.Replace(Environment.NewLine,"").Trim();
			return vysledek == "1";*/
 			
 			 			
 			string info = m_services.RequestInformation("MessagesDisplayed");
 			return true;
		}
		

		
		
		#region IDisposable implementation
		public void Dispose()
		{

			
			if (_dialogy) {
				m_services.DoCommand(m_token,"DIALOGS MESSAGE ON");
				m_services.DoCommand(m_token,"DIALOGS ERROR ON");
			}
			


		}
		#endregion
	}
}
