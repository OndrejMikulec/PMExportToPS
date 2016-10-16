/*
 * Created by SharpDevelop.
 * User: omiku
 * Date: 16.10.2016
 * Time: 8:57
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace PMExportToPS
{
	/// <summary>
	/// Description of DoWork.
	/// </summary>
	public class DoWork
	{
		object _PSApplication = null;
		object _PSDocument = null;
		
		string _pSPath;
		
		public DoWork(string m_token, PowerMILL.PluginServices m_services, string _pSPath)
		{
			this._pSPath = _pSPath;
			
			bool selectedSurfaces = false;
			object oSelectedSurfaces;
			m_services.DoCommandEx(m_token,"print selsurface",out oSelectedSurfaces);
			
			if (!string.IsNullOrEmpty(oSelectedSurfaces.ToString())) {
				selectedSurfaces = true;
			}
			
			m_services.DoCommand(m_token,@"STRING $modelNameForApp = """"");
			m_services.DoCommand(m_token,@"$modelNameForApp = """"");
			m_services.DoCommand(m_token,@"$modelNameForApp =  INPUT ENTITY MODEL ""Vyber model pro export""");
			
			object modelName;
			m_services.DoCommandEx(m_token,"print $modelNameForApp",out modelName);
			
			if (!string.IsNullOrEmpty(modelName.ToString())) {
				if (!CreateConection()) {
					if (!startPSandConnect()) {
						System.Windows.Forms.MessageBox.Show("Connection to PS failed");
					}
				}
				
				object modelPath;
				m_services.DoCommandEx(m_token,@"print $entity(""Model"","""+modelName+@""").Path", out modelPath);
				
				if (!File.Exists(modelPath.ToString())||selectedSurfaces) {
					modelPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)+@"\"+modelName+".dgk";
					m_services.DoCommand(m_token,@"EXPORT MODEL """+modelName+@""" FILESAVE """+modelPath+@"""");
					ImportModel(modelPath.ToString());
					File.Delete(modelPath.ToString());
				} else {
					ImportModel(modelPath.ToString());
				}
			}
			
			
		}
		
		void DoCommand(string com)
		{
			Console.WriteLine(com);
			NewLateBinding.LateCall(_PSApplication, (Type)null, "Exec", new object[] { com }, (string[])null, (Type[])null, new bool[] { true }, true);
			
		}
		
        bool CreateConection()
        {
            if (_PSApplication == null)
            {
                try
                {
                    _PSApplication = RuntimeHelpers.GetObjectValue(Interaction.GetObject((string)null, "PowerSHAPE.Application"));
                    _PSDocument = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(_PSApplication, (Type)null, "ActiveDocument", new object[0], (string[])null, (Type[])null, (bool[])null));
                }
                catch
                {
                    return false;
                }

                return true;
            }
            else
                return false;
        }
        
        bool startPSandConnect()
		{
			if (File.Exists(_pSPath)) {
				Process.Start(_pSPath);
				int count = 30;
				while (count>0) {
					Thread.Sleep(1000);
					count--;
					if (!CreateConection()) {
						break;
					} else {
						return true;
					}
				}
			}
        	
        	return false;
		}
        
         void ImportModel(string Filename)
		{
			SetDialogMode(false);
			DoCommand("PRINT CREATED.CLEARLIST");
			DoCommand("FILE IMPORT '" + Filename + "'");
			DoCommand("DISMISS");
			
			SetDialogMode(true);
			
		}
        
        void SetDialogMode(bool DialogOn)
		{
			if (DialogOn) {
        		DoCommand("DIALOG ON");
			}
			else {
				DoCommand("DIALOG OFF");
			}
		}
		
		
	}
}
