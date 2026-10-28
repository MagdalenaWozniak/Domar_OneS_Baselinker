using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using cdn_api;

namespace OneS_Baselinker
{
    public class Configuration
    {
        public int WersjaAPI { set; get; }
        public int Sesja { set; get; }
        public string XLLogin_Baza { set; get; }
        public string XLLogin_OpeIdent { set; get; }
        public string XLLogin_OpeHaslo { set; get; }
        public string XLLogin_ProgramID { set; get; }
        public int XLLogin_TrybWsadowy { set; get; }
        public int XLLogin_Winieta { set; get; }
        public int XLLogin_UtworzWlasnaSesje { set; get; }
        public int BLTest { set; get; }
        public int BLDays { set; get; }
        public string XLServer { set; get; }
        public string XLDatabase { set; get; }
        public string XLConString { set; get; }
        

        public string ErrorDescription(int funkcja, int blad)
        {
            XLKomunikatInfo_20220 KomunikatInfo20220 = new XLKomunikatInfo_20220
            {
                Wersja = this.WersjaAPI,
                Funkcja = funkcja,
                Blad = blad
            };
            cdn_api.cdn_api.XLOpisBledu(KomunikatInfo20220);
            if (KomunikatInfo20220.OpisBledu != null)
                _ = (string)KomunikatInfo20220.OpisBledu;
            return (string)KomunikatInfo20220.OpisBledu;
        }
    }
}
