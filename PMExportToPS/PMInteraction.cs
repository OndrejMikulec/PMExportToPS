/*
 * Created by SharpDevelop.
 * User: Ondra
 * Date: 25/03/2016
 * Time: 16:15
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Windows.Forms;

namespace PMExportToPS
{
	/// <summary>
	/// Description of PMInteraction.
	/// </summary>
	public class PMInteraction : IDisposable
	{
		
		bool _dialogy;
		bool _error;
		
		string m_token;
		PowerMILL.PluginServices m_services;
		
		public PMInteraction(string m_token,PowerMILL.PluginServices m_services)
		{
			this.m_token = m_token;
			this.m_services = m_services;
			

			
			m_services.QueueCommand(m_token,@"ECHO OFF DCPDEBUG UNTRACE COMMAND ACCEPT");
			
			_dialogy = m_services.RequestInformation("MessagesDisplayed") == "true";
			_error = m_services.RequestInformation("ErrorsDisplayed") == "true";
			
			m_services.QueueCommand(m_token,"DIALOGS MESSAGE OFF");
			m_services.QueueCommand(m_token,"DIALOGS ERROR OFF");

		}
		

		
		
		#region IDisposable implementation
		public void Dispose()
		{
			m_services.QueueCommand(m_token, "QUIT");
			m_services.QueueCommand(m_token, "QUIT");
			m_services.QueueCommand(m_token, "QUIT");
			
			if (_dialogy)
				m_services.QueueCommand(m_token, "DIALOGS MESSAGE ON");
			if (_error)
				m_services.QueueCommand(m_token, "DIALOGS ERROR ON");


		}
		#endregion
	}
}
