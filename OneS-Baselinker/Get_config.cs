using System;
using System.IO;
using System.Xml.Serialization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace OneS_Baselinker
{
    public static class Get_config
    {
        public static Configuration GetData()
        {
            Configuration configuration = new Configuration();
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Configuration));
            FileStream fileStream = (FileStream)null;
            try
            {
                fileStream = new FileStream(@"C:\OneSupport\Baselinker\Baselinker_Config.xml", FileMode.Open);
                configuration = (Configuration)xmlSerializer.Deserialize((Stream)fileStream);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                if (fileStream != null)
                    fileStream.Close();
            }
            return configuration;
        }
    }
}
