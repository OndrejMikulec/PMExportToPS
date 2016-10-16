/*
 * Created by SharpDevelop.
 * User: val01039
 * Date: 17.2.2016
 * Time: 6:10
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Globalization;

namespace PMExportToPS
{

	public static class Gets
	{
		
		public static double getCultureInvariantDouble(string str)
		{
			return  double.Parse(str.Trim().Replace(",","."),CultureInfo.InvariantCulture);
		}
				
		/*public static string getMainWorkplane(PowerMILL.Application oAppliacation)
		{
			List<string> jmenaWorkplanes = Gets.getEntitySorting("Workplane",oAppliacation);
			if (jmenaWorkplanes.Count > 0)
				return jmenaWorkplanes[0];
			return null;

		}
		
		public static string getNamedEntityParameterString(string entity,string jmeno, string parameter,PowerMILL.Application oAppliacation)
		{
			
 			int err = 0;
 			string returnValue = null;
 			oAppliacation.ExecuteEx(@"print $entity("""+entity+@""","""+jmeno+@""")."+parameter,out err,out returnValue);
 			returnValue = returnValue.Replace(Environment.NewLine,"").Trim();
 			
 			return returnValue;
		}
		
		public static bool getNamedEntityParameterBool(string entity,string jmeno, string parameter,PowerMILL.Application oAppliacation)
		{
			string str = Gets.getNamedEntityParameterString(entity,jmeno,parameter,oAppliacation);
			return str.Trim() == "1"; 
		}
		
		public static int getNamedEntityParameterInt(string entity,string jmeno, string parameter,PowerMILL.Application oAppliacation)
		{
			string str = Gets.getNamedEntityParameterString(entity,jmeno,parameter,oAppliacation);
			return  int.Parse(str.Trim());
		}
		
		public static double getNamedEntityParameterDouble(string entity,string jmeno, string parameter,PowerMILL.Application oAppliacation)
		{
			string str = Gets.getNamedEntityParameterString(entity,jmeno,parameter,oAppliacation);
			return  double.Parse(str.Trim());
		}
		
		public static double[] getNamedEntityParameterVector3(string entity,string jmeno, string parameter,PowerMILL.Application oAppliacation)
		{
			double d0 = Gets.getNamedEntityParameterDouble(entity,jmeno,parameter+"[0]",oAppliacation);
			double d1 = Gets.getNamedEntityParameterDouble(entity,jmeno,parameter+"[1]",oAppliacation);
			double d2 = Gets.getNamedEntityParameterDouble(entity,jmeno,parameter+"[2]",oAppliacation);
			
			return  new double[] {d0,d1,d2 };
				
		}*/
		
		public static string getActiveEntityParameterString(string entity, string parameter,string m_token,PowerMILL.PluginServices m_services)
		{

			
 			object returnValueO = null;
 			m_services.DoCommandEx(m_token,@"print $entity("""+entity+@""","""")."+parameter,out returnValueO);
 			string returnValue = returnValueO.ToString().Replace(Environment.NewLine,"");
 			
 			return returnValue;
		}
		
		public static bool getActiveEntityParameterBool(string entity,string parameter,string m_token,PowerMILL.PluginServices m_services)
		{
			string str = Gets.getActiveEntityParameterString(entity,parameter,m_token,m_services);
			return str.Trim() == "1"; 
		}
		
		public static int getActiveEntityParameterInt(string entity,string parameter,string m_token,PowerMILL.PluginServices m_services)
		{
			string str = Gets.getActiveEntityParameterString(entity,parameter,m_token,m_services);
			return  int.Parse(str.Trim());
		}
		
		public static double getActiveEntityParameterDouble(string entity, string parameter,string m_token,PowerMILL.PluginServices m_services)
		{
			string str = Gets.getActiveEntityParameterString(entity,parameter,m_token,m_services);
			return  getCultureInvariantDouble(str);

		}
		
		public static double[] getActiveEntityParameterVector3(string entity, string parameter,string m_token,PowerMILL.PluginServices m_services)
		{
			double d0 = Gets.getActiveEntityParameterDouble(entity,parameter+"[0]",m_token,m_services);
			double d1 = Gets.getActiveEntityParameterDouble(entity,parameter+"[1]",m_token,m_services);
			double d2 = Gets.getActiveEntityParameterDouble(entity,parameter+"[2]",m_token,m_services);
			
			return  new double[] {d0,d1,d2 };
				
		}
		
		public static string getGlobalParameterString(string parameter,string m_token,PowerMILL.PluginServices m_services)
		{
		
 			object returnValueO = null;
			m_services.DoCommandEx(m_token,@"print $"+parameter,out returnValueO);
 			string returnValue = returnValueO.ToString().Replace(Environment.NewLine,"");
 			return returnValue;
		}
		
		public static bool getGlobalParameterBool(string parameter,string m_token,PowerMILL.PluginServices m_services)
		{
			string str = Gets.getGlobalParameterString(parameter,m_token,m_services);
			return str.Trim() == "1"; 
		}
		
		public static int getGlobalParameterInt(string parameter,string m_token,PowerMILL.PluginServices m_services)
		{
			string str = Gets.getGlobalParameterString(parameter,m_token,m_services);
			return  int.Parse(str.Trim());
		}
		
		public static double getGlobalParameterDouble(string parameter,string m_token,PowerMILL.PluginServices m_services)
		{
			string str = Gets.getGlobalParameterString(parameter,m_token,m_services);
			return  getCultureInvariantDouble(str);
			
		}
		
		public static double[] getGlobalParameterVector3(string parameter,string m_token,PowerMILL.PluginServices m_services)
		{
			double d0 = Gets.getGlobalParameterDouble(parameter+"[0]",m_token,m_services);
			double d1 = Gets.getGlobalParameterDouble(parameter+"[1]",m_token,m_services);
			double d2 = Gets.getGlobalParameterDouble(parameter+"[2]",m_token,m_services);
			
			return  new double[] {d0,d1,d2 };
				
		}
		
		
		public static List<string> getEntitySorting(string entityType,string m_token,PowerMILL.PluginServices m_services )
		{

			
			List<string> returnValue = null;
			try {
				
				string pathMacroTemp = getCestaMakroTemp();

				using (var writeMacro = new StreamWriter(pathMacroTemp,false,Encoding.Default)) {
					writeMacro.WriteLine(@"FOREACH $x IN FOLDER("""+entityType+@""") {");
					writeMacro.WriteLine("print $x.Name");
					writeMacro.WriteLine("}");
				}
				
	 			object resultO = null;
	 			m_services.DoCommandEx(m_token,@"macro """+pathMacroTemp+@"""",out resultO);
	 			string result = resultO.ToString().Replace(Environment.NewLine,"@");
	 			string[] resultList = result.Split('@');
	 			returnValue = new List<string>();
	 			foreach (string ft in resultList) {
					if (!string.IsNullOrEmpty(ft)) {
						returnValue.Add(ft);
					}
	 			}
	 			
	 			
			} catch{ returnValue = null; System.Windows.Forms.MessageBox.Show("Chyba čtení temp makra!");}
			
			return returnValue;
		}
		
		
		public static string cestaSlozkyExeSouboru()
		{
			string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        	return Path.GetDirectoryName(exePath);
		}
		
		public static string  getCestaMakroTemp()
		{
        	return cestaSlozkyExeSouboru() + @"\MakroTemp.mac";
		}
		
		public static string  getCestahelpSouboru()
		{
        	return cestaSlozkyExeSouboru() + @"\Help.pdf";
		}
		
		
		public static string  getExportModelTemp()
		{
			string cestaumisteni = cestaSlozkyExeSouboru();
			
        	string returnValue = cestaumisteni + @"\temp.dmt";
        	
        	int count = 1;
        	while (File.Exists(returnValue)) {
        		returnValue = cestaumisteni + @"\temp"+count+".dmt";
        		count++;
        	}
        	
        	return returnValue;
		}
		
		public static string getMainModel(string m_token,PowerMILL.PluginServices m_services)
		{
			
			List<string> modelList = getEntitySorting("Model", m_token,m_services);
			
			if (modelList.Count>0) {
				return modelList[0];
			}
			return null;
		}

	}
}
