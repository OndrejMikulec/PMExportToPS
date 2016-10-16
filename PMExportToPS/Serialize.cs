/*
 * Created by SharpDevelop.
 * User: Ondra
 * Date: 28/02/2016
 * Time: 14:23
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Drawing;
using System.IO;
using System.Runtime.Serialization;
using System.Windows.Forms;


namespace PMExportToPS
{
	/// <summary>
	/// Description of MainFormSerialize.
	/// </summary>
	 [DataContract] 
	 public class MySerialization
	{
		[DataMember] string _psPath;
		
		
		
		MySerialization(Plugin plg)
		{
			_psPath = plg.PathPS;
			
		}
		
		
		public static void Save(Plugin plg)
		{
			MySerialization oMainFormSerialize = new MySerialization(plg);
			
			NetDataContractSerializer dcs = new NetDataContractSerializer();
			
			using (Stream oStream = File.Create(plg.SerializePath)) {
				dcs.WriteObject(oStream,oMainFormSerialize);
			}
		}
		
		public static void Load(Plugin plg)
		{

			if (!File.Exists(plg.SerializePath)) {
				setdefaultVal(plg);
				return;
			}
			
			try {
			
				NetDataContractSerializer dcs = new NetDataContractSerializer();
				
				MySerialization oMainFormSerialize = null;
				using (Stream oStream = File.Open(plg.SerializePath,FileMode.Open,FileAccess.Read)) {
					oMainFormSerialize = (MySerialization)dcs.ReadObject(oStream);
				}
				plg.PathPS = oMainFormSerialize._psPath;
				
			
			} catch (Exception) {
				
				setdefaultVal(plg);
			}
			
			
		}
		
		static void setdefaultVal(Plugin plg)
		{
            
		}


	}
}
