/*
 * Created by SharpDevelop.
 * User: omiku
 * Date: 16.10.2016
 * Time: 8:59
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace PMExportToPS
{
	/// <summary>
	/// Description of Options.
	/// </summary>
	public partial class Options : Form
	{
		string m_token; 
		PowerMILL.PluginServices m_services;
		Plugin _plg;
		
		public Options(string m_token, PowerMILL.PluginServices m_services , Plugin plg )
		{
			this.m_token = m_token;
			this.m_services = m_services;
			_plg = plg;
			
			InitializeComponent();
			
			label1.Text = plg.PathPS;
			label2.Text = plg.SerializePath;
		}
		void Button1Click(object sender, EventArgs e)
		{
			OpenFileDialog fl = new OpenFileDialog() {};//TODO
			
			if (fl.ShowDialog()==DialogResult.OK) {
				_plg.PathPS = fl.FileName;
				label1.Text = _plg.PathPS;
			}
		}
		void Button2Click(object sender, EventArgs e)
		{
			Close();
		}
	}
}
