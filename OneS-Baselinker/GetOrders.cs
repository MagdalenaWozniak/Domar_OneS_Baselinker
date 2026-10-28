using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.IO;
using System.ComponentModel;
using cdn_api;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Threading;
using System.Collections;

namespace OneS_Baselinker
{
    class GetOrders
    {
        //DEBUG
        public static int dev = 0;

        //XML Config
        private Configuration cnf = new Configuration();
        public static GetOrders program = new GetOrders();
        //XL Dane
        private  string XLProgramID = "GetOrders";
        private string XLDatabase = "DOMAR2"; // "DomarXL";  //"DOMAR2"; "KopiaXL"
        private string XLUser = "OS_AUTOM";        //"ADMIN"; "OS_AUTOM"
        private string XLPass = "DomXL#53&"; //"Dupa2003!";
        private string XLKey = "SRV1::5000139268"; //"localhost::5000124544";//"SRV1::5000139268"; //localhost::5000124544 albo "localhost::5000139271" ??
        private string XLConString = "; User ID=ADMIN;Password=Dupa2003!"; //"; User ID = SA; Password=DoMarSQL17"; //"; User ID=ADMIN;Password=Dupa2003!";//"; User ID=SA;Password=DoMarSQL17";

        private int test =0;
        private int days = 7;
        private  int XLWersjaAPI = 20220;
        private  int XLTrybWsadowy = 1;
        private  int XLWinieta = -1;
        private  int XLUtworzWlasnaSesje =1;

        //Odpowiedź XL API
        private  int XLdokument_wynik = -1; 
        private  int XLzamknij_dokument_wynik = -1;
        private  int XLwydrukuj_dokument_wynik = -1;
        private  int XLdokument_pozycje_wynik = -1;
        private  int XLdokument_platnosc_Wynik = -1;
        private  static int XLdokument_id = -1;
        private  static int XLFaktura_id = -1;
        
        private  static int XLkontrahent_id = -1;
        private  static int XLadres_id = -1;
        private  static int XLtowar_id = 0;
        //XLSQL 
        private  string XLSQLServer;
        private  string XLSQLDatabase;

        //XL Things
        private XLDokumentZamElemInfo_20220 xlDokumentZamElemInfo;
        private XLDokumentZamNagInfo_20220 xlDokumentZamInfo;
        private XLDokumentNagInfo_20220 dokument;
        private XLTowarInfo_20220 xLTowarInfo;
        private XLCenaInfo_20220 xLCenaInfo_20220;
        private XLDokumentElemInfo_20220 dokument_pozycje;
        private XLKontrahentInfo_20220 kontrahentInfo;
        private XLPlatnoscInfo_20220 dokument_platnosc;
        private XLZamkniecieDokumentuInfo_20220 zamknij_dokument;
        private XLWydrukInfo_20220 wydrukuj_dokument;
        private XLKomunikatInfo_20220 info;

        //duperele
        private static Baselinker_response cumulative_response;
        private static Baselinker_journal_response journal_response;
        private static int[] KlasyAtrybutu = new int[13] {2, 31, 42, 34, 32, 36, 43, 44, 40, 41, 37, 33, 65};
        private static string[] KlasyAtrybutuNazwa = new string[13] { "Dealer", "Status Zamówienia", "Czy faktura", "ID BL", "ID Zamówienia", "Źródło", "Link do zamówienia", "Login", "BL Sposób dostawy", "BL Sposób Płatności", "Status Płatności", "Id płatności", "Przewoźnik" };
        private static string cumulative_json = "{\"status\":\"SUCCESS\",\"orders\":[";

        public static string connection_string = "";
        private static int sesja = -1;
        private static int[] dokInfo;

        [DllImport("ClaRUN.dll")]
        public static extern void AttachThreadToClarion(bool x);

        private string RemoveExtraText(string value)
        {
            var allowedChars = "01234567890.,";
            return new string(value.Where(c => allowedChars.Contains(c)).ToArray());
        }
        public static DateTime UnixTimeStampToDateTime(double unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            System.DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            dtDateTime = dtDateTime.AddSeconds(unixTimeStamp).ToLocalTime();
            return dtDateTime;
        }

        public static string StripText(string title)
        {
            string output = title.Trim();
            string[,] znakiSpecjalne = {
            { "Ą", "A" }, { "Ć", "C" }, { "Ę", "E" }, { "Ń", "N" }, { "Ó", "O" }, { "Ś", "S" }, { "Ź", "Z" }, { "Ż", "Z" },
            { "ą", "a" }, { "ć", "c" }, { "ę", "e" }, { "ń", "n" }, { "ó", "o" }, { "ś", "s" }, { "ź", "z" }, { "ż", "z" },
            };
            for (int i = 0; i < znakiSpecjalne.GetLength(0); i++)
            {
                output = output.Replace(znakiSpecjalne[i, 0], znakiSpecjalne[i, 1]);
            }

            return output;
        }
        public static string Get_time(int days)
        {
            TimeSpan t = DateTime.UtcNow - new DateTime(1970, 1, 1);
            int secondsSinceEpoch = (int)t.TotalSeconds - (3600 * 24 * days);
            return Convert.ToString(secondsSinceEpoch);
        }
        public static int GetClarionDate(DateTime dotNetDate)
        {
            CultureInfo ci = new CultureInfo("pl-PL");
            DateTime baseDate = DateTime.Parse("01.01.1801", ci);
            System.TimeSpan ts = dotNetDate - baseDate;
            int l = ts.Days + 4;
            return l;
        }
        public static int GetClarionTime(DateTime time)
        {
            int l = time.Hour * 360000;
            l += time.Minute * 6000;
            l += time.Second * 100;
            l += time.Millisecond / 10;
            return l + 1;
        }
        public string Get_Version()
        {
            return typeof(String).Assembly.GetName().Version.ToString();
        }
        public bool Get_Config()
        {
            bool flaga = true;


            try
            {
                cnf = Get_config.GetData();
                XLUser = cnf.XLLogin_OpeIdent;
                XLDatabase = cnf.XLLogin_Baza;
                XLPass = cnf.XLLogin_OpeHaslo;
                XLProgramID = cnf.XLLogin_ProgramID;
                XLWersjaAPI = 20220;//cnf.WersjaAPI;
                XLKey = cnf.XLServer;
                XLConString = cnf.XLConString;
                test = cnf.BLTest;
                days = cnf.BLDays;
                
            }
            catch (Exception e)
            {
                Console.WriteLine("Błąd get_config: "+ e.Message);
                //Console.ReadKey();
                flaga = false;
            }

            //XLTrybWsadowy = cnf.XLLogin_TrybWsadowy;
            //XLWinieta = cnf.XLLogin_Winieta;
            //XLUtworzWlasnaSesje = -1;

            //XLSQLServer = cnf.XLServer;
            //XLSQLDatabase = cnf.XLDatabase;
            //XLSQLServer = "192.168.137.149";
            //XLSQLDatabase = "DOMAR";

            return flaga;                
        }
        public int Login_XL()
        {
            //AttachThreadToClarion(true);
            Console.WriteLine("XLProgramID");
            return cdn_api.cdn_api.XLLogin(new XLLoginInfo_20220()
            {
                Wersja = 20220,//XLWersjaAPI,
                UtworzWlasnaSesje = XLUtworzWlasnaSesje,
                Winieta = XLWinieta,
                TrybWsadowy = XLTrybWsadowy,
                ProgramID = XLProgramID,
                Baza = XLDatabase, //TestXL 
                OpeIdent = "OS_AUTOM", //XLUser, //"OS_AUTOM" //OS
                OpeHaslo = "DomXL#53&", //XLPass, //"DomXL#53&"//NE!w0rk12, 
                TrybNaprawy = 1,
                //PlikLog = (@"C:\OneSupport\Baselinker\api.log"),
                SerwerKlucza =XLKey,//"localhost::5000139271",//"SRV1::5000139268", 
            }, ref sesja);
            
        }
        public static void Logout_XL()
        {
            AttachThreadToClarion(true);
            string str = "";
            XLLogoutInfo_20220 xLLogoutInfo_20220 = new XLLogoutInfo_20220()
            {
                Wersja = 20220,
                //TrybUsuwaniaDok = 1,
            };
            str = cdn_api.cdn_api.XLLogoutEx(xLLogoutInfo_20220, sesja).ToString();
            Console.WriteLine(str);

        }

        //Connection String
        public string GetConnectionString()
        {
            //XLPolaczenieInfo_20220 polaczenieInfo = new XLPolaczenieInfo_20220 { Wersja = 20220 };
            //cdn_api.cdn_api.XLPolaczenie(polaczenieInfo);
            XLPolaczenieInfo_20220 polaczenieInfo = new XLPolaczenieInfo_20220 { Wersja = 20220 };
            cdn_api.cdn_api.XLPolaczenie(polaczenieInfo);

            //connection_string = @"Provider=MSOLEDBSQL.1;Password=43gQdfs31so3;Persist Security Info=True;User ID=onesupport_dod;Initial Catalog="";Data Source=SRV1;Initial File Name="";Server SPN="";Authentication="";Access Token=""";
            connection_string = "Server=" + polaczenieInfo.Serwer + ";Database=" + polaczenieInfo.Baza + XLConString;//"; User ID=SA;Password=DoMarSQL17";//";Integrated Security=SSPI;";//"; User ID=SA;Password=DoMarSQL17"; // "; User ID=SA;Password=DoMarSQL17"; // ";Integrated Security=SSPI;";
            Console.WriteLine(connection_string);
            return connection_string;
        }
        public static string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }
        public List<string> ConnectListToSQL(string connection_string, string query)
        {
            List<string> columnData = new List<string>();

            using (SqlConnection connection = new SqlConnection(connection_string))
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = new SqlCommand(query, connection);
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet, "Query");
                if (dataSet.Tables.Count == 0)
                {
                    return columnData;
                }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                    for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                    {
                    
                        columnData.Add(row[0].ToString());
                    }
                }
            }
            return columnData;
        }

        static public async Task<Baselinker_response> Http_Baselinker_Response(InputParameters methodParams, bool SaveToFile, string method)
        {
            var client = new HttpClient();
            string json = "";

            Console.WriteLine(methodParams.date_confirmed_from);

            var apiParams = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("token", "4150-22932-JOD6ZGJ2FQ7LEIR5LNHJ0KWX3JNG9BEKJYCHTBZ0YX6M3Z9J4BT23S4TBGSEPB86"),
                new KeyValuePair<string, string>("method", method),
                new KeyValuePair<string, string>("parameters", JsonConvert.SerializeObject(methodParams)),
            });

            HttpResponseMessage response = await client.PostAsync(
                "https://api.baselinker.com/connector.php",
                apiParams);
            HttpContent responseContent = response.Content;
            var baselinker_response = new Baselinker_response();
            using (var reader = new StreamReader(await responseContent.ReadAsStreamAsync()))
            {
                if (dev == 0)
                {
                    json = await reader.ReadToEndAsync();
                    
                    baselinker_response = Baselinker_response.FromJson(json);
                    if(SaveToFile == true)
                    {
                        TextWriter tw2 = new StreamWriter(@"C:\OneSupport\Baselinker\Input.txt");
                        tw2.WriteLine(json);
                        tw2.Close();
                    }
                    try
                    {
                        json = json.Substring(30);
                        cumulative_json += json;
                        cumulative_json = cumulative_json.Remove(cumulative_json.Length - 2);
                        cumulative_json += ",";
                    }
                    catch(Exception e)
                    {
                        System.Environment.Exit(-1);
                    }
                }
                else
                {
                    json = File.ReadAllText(@"C:\OneSupport\Baselinker\Input.txt");
                    if (json != "" && json != null) baselinker_response = Baselinker_response.FromJson(json);
                    else System.Environment.Exit(-1);
                }
            }
            client.Dispose();
            return baselinker_response;
        }
        static public async Task<Baselinker_journal_response> Journal_Http_Baselinker_Response(JournalParameters methodParams, bool SaveToFile, int numbers, string method)
        {
            var client = new HttpClient();
            string json = "";

            var apiParams = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("token", "4150-22932-JOD6ZGJ2FQ7LEIR5LNHJ0KWX3JNG9BEKJYCHTBZ0YX6M3Z9J4BT23S4TBGSEPB86"),
                new KeyValuePair<string, string>("method", method),
                new KeyValuePair<string, string>("parameters", JsonConvert.SerializeObject(methodParams)),
            });

            HttpResponseMessage response = await client.PostAsync(
                "https://api.baselinker.com/connector.php",
                apiParams);
            HttpContent responseContent = response.Content;
            var baselinker_response = new Baselinker_journal_response();
            using (var reader = new StreamReader(await responseContent.ReadAsStreamAsync()))
            {
                if (dev == 0)
                {
                    json = await reader.ReadToEndAsync();

                    baselinker_response = Baselinker_journal_response.FromJson(json);
                    if (SaveToFile == true)
                    {
                        TextWriter tw2 = new StreamWriter(@"C:\OneSupport\Baselinker\Journal.txt");
                        tw2.WriteLine(json);
                        tw2.Close();
                    }
                    try
                    {
                        json = json.Substring(numbers);
                        cumulative_json += json;
                        cumulative_json = cumulative_json.Remove(cumulative_json.Length - 2);
                        cumulative_json += ",";
                    }
                    catch (Exception e)
                    {
                        System.Environment.Exit(-1);
                    }
                }
                else
                {
                    json = File.ReadAllText(@"C:\OneSupport\Baselinker\Journal.txt");
                    if (json != "" && json != null) baselinker_response = Baselinker_journal_response.FromJson(json);
                    else System.Environment.Exit(-1);
                }
            }
            client.Dispose();
            return baselinker_response;
        }
        static public async Task Baselinker_Connect(int days)
        {
            InputParameters methodParams = new InputParameters
            {
                date_confirmed_from = Get_time(days)
            };
            Baselinker_response baselinker_response = new Baselinker_response();
            TextWriter tw = new StreamWriter(@"C:\OneSupport\Baselinker\WriteLinesFromText.txt");
            int count = 0;
            do
            {
                baselinker_response = await Http_Baselinker_Response(methodParams, true,"getOrders");
                if (baselinker_response.Orders.Length == 0) throw new Exception("Brak zamówień z wybranego okresu!");
                for (int i = 0; i < baselinker_response.Orders.Length; i++)
                {
                    try
                    {
                        foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(baselinker_response.Orders[i]))
                        {
                            string name = descriptor.Name;
                            object value = descriptor.GetValue(baselinker_response.Orders[i]);
                            if (name == "Products")
                            {
                                //tw.WriteLine(name + " : ");
                                //Console.WriteLine(name + " : ");
                                for (int j = 0; j < baselinker_response.Orders[i].Products.Length; j++)
                                {
                                    foreach (PropertyDescriptor descriptor2 in TypeDescriptor.GetProperties(baselinker_response.Orders[i].Products[j]))
                                    {
                                        object value_2 = descriptor2.Name;
                                        object product_value = descriptor2.GetValue(baselinker_response.Orders[i].Products[j]);
                                        //tw.WriteLine(value_2 + "=" + product_value);
                                        //Console.WriteLine("{0}={1}", value_2, product_value);
                                    }
                                }
                            }
                            else
                            {
                                //tw.WriteLine(name + "==" + value);
                            }
                            if (name == "DateConfirmed")
                            {
                                Console.WriteLine("{0} = {1}", name, value);
                            }
                        }
                        Console.WriteLine("----" + (i + 1 + count * 100) + "----");
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine("Błąd pobierania zamówienia: " + baselinker_response.Orders[i].OrderId + "Błąd: " + e.Message);
                    }
                }
                methodParams.date_confirmed_from = baselinker_response.Orders[baselinker_response.Orders.Length-1].DateConfirmed;
                //baselinker_response = await Http_Baselinker_Response(methodParams, false);
                count++;

            } while (baselinker_response.Orders.Length == 100);
                // close the stream
            tw.Close();
            cumulative_json = cumulative_json.Remove(cumulative_json.Length - 1);
            cumulative_json += "]}";

            TextWriter tw2 = new StreamWriter(@"C:\OneSupport\Baselinker\Input.txt");
            tw2.WriteLine(cumulative_json);
            tw2.Close();
            cumulative_response = Baselinker_response.FromJson(cumulative_json);
            Console.WriteLine("Zakończono pobieranie z API, ilosc cumulative to " + cumulative_response.Orders.Length);
        }
        static public async Task<int> AnalyseJournal(string order_id, string status,string form, string[] logs_types, TextWriter tw_log)
        {
            JournalParameters methodParams = new JournalParameters
            {
                logs_types=logs_types,
                order_id=order_id
            };
            Baselinker_journal_response baselinker_Journal_Response = new Baselinker_journal_response();
            baselinker_Journal_Response = await Journal_Http_Baselinker_Response(methodParams, true,28, "getJournalList");
            if (baselinker_Journal_Response.Logs.Length == 0) return 0;

            string resp = "";
            for (int i = 0; i < baselinker_Journal_Response.Logs.Length; i++)
            {
                if (i != 0)
                {
                    resp += ";";
                }
                resp += baselinker_Journal_Response.Logs[i].ObjectId;
            }
            for (int i = 0; i < baselinker_Journal_Response.Logs.Length; i++)
            {
               
                if ((baselinker_Journal_Response.Logs[i].ObjectId == "237649" || baselinker_Journal_Response.Logs[i].ObjectId == "208277" || baselinker_Journal_Response.Logs[i].ObjectId == "218865") && resp.Contains(status.Trim()))   //.Contains(status))
                {
                    Console.WriteLine(baselinker_Journal_Response.Logs[i].ObjectId);
                    tw_log.WriteLine($@"Odnaleziono wymagany status dla zamówienia {order_id}. Aktualne statusy to: {resp}, status pobranego zamówienienia to: {status}");
                    return 1;
                }       
            }
            Console.WriteLine($@"Nie odnaleziono wymaganego statusu dla zamówienia {order_id}. Aktualne statusy to: {resp}, status pobranego zamówienienia to: {status}");
            tw_log.WriteLine($@"Nie odnaleziono wymaganego statusu dla zamówienia {order_id}. Aktualne statusy to: {resp}, status pobranego zamówienienia to: {status}");
            //tw_log.WriteLine("resp contain status: " + resp.Contains(status.Trim()));
            return 0;
            
        }
        private static async Task CreateIndex_XL()
        {
            AttachThreadToClarion(true);
            //Dodawanie nowego zamówienia
            string query = "";
            int[] response = { 32, 1034247, 0, 0 };
            TextWriter tw_log = new StreamWriter(@"C:\OneSupport\Baselinker\log.txt", append: true);
            
            for (int i = 0; i < cumulative_response.Orders.Length; i++)
            {

                try
                {
                    //tw_log.WriteLine("Ilosc zamowien to : " + cumulative_response.Orders.Length);
                    //Czy istnieje takie zamowienie w XL

                    query = "SELECT TOP 1 1 from CDN.ZamNag where Zan_DokumentObcy = 'ESKLEP " + cumulative_response.Orders[i].OrderId + "'";
                    if (ConnectToSQL(connection_string, query) == "1")
                    {
                        Console.WriteLine("Zamowienie " + i + " o orderID " + cumulative_response.Orders[i].OrderId + " jest już dodane.");
                    }

                    else if (await AnalyseJournal(cumulative_response.Orders[i].OrderId, cumulative_response.Orders[i].OrderStatusId, "Cod", new string[] { "18" }, tw_log) == 1)
                    {

                        
                        Console.WriteLine(ConnectToSQL(connection_string, "select name_for_customer from ord.BL_StatusID where id=" + cumulative_response.Orders[i].OrderStatusId));
                        tw_log.WriteLine("Status po zaakceptowaniu: " + ConnectToSQL(connection_string, "select name_for_customer from ord.BL_StatusID where id=" + cumulative_response.Orders[i].OrderStatusId));
                        //Dodawanie kontrahenta z nipem lub bez
                        if (string.IsNullOrEmpty(cumulative_response.Orders[i].InvoiceNip))
                        {
                            System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(@"\D");
                            cumulative_response.Orders[i].InvoiceNip = rgx.Replace(cumulative_response.Orders[i].InvoiceNip, "");
                        }
                        else
                        {
                            cumulative_response.Orders[i].InvoiceNip = Regex.Replace(cumulative_response.Orders[i].InvoiceNip, @"[^\d]", "");
                        }
                        //Console.WriteLine(cumulative_response.Orders[i].InvoiceNip);
                        //Console.WriteLine("Nie istnieje kontrahent o takim NIP");
                        response = program.AddContractor_XL(i, tw_log);
                        if (response[2] < 1)
                        {
                            continue;
                        }
                        Console.WriteLine("1. Dodawanie nowego kontrahenta bez NIP o nowym GID " + response[2]);
                        tw_log.WriteLine(DateTime.Now + "1. GID Nowego kontrahenta to " + response[2]);

                        //Dodaj ZS
                        try {
                            Console.WriteLine("GID TYP: " + response[0] + "\nGID Firma: " + response[1] + "\nGID Numer: " + response[2] + " \nGID Lp: " + response[3]);
                            Console.WriteLine("i: "+ i);
                            Console.WriteLine("Order id: "+cumulative_response.Orders[i].OrderId);
                            Console.WriteLine("Tw_log: " + tw_log.GetType() );
                            program.AddZSOrder_XL(i, response, cumulative_response.Orders[i].OrderId, tw_log);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Wystąpił błąd w dodawaniu ZS " + ex.ToString() + "\nDla order id: " + cumulative_response.Orders[i].OrderId);
                        }

                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Wystąpił błąd w CreateIndex " + e.ToString() + "\nDla order id: "+ cumulative_response.Orders[i].OrderId);
                }
            }
            tw_log.WriteLine("Sprawdzanie, czy są FS do wystawienia");
            program.AddFS_FROM_ZS(tw_log);
            tw_log.WriteLine("----------");
            tw_log.Close();
        }
        public void SplitDeliveryName(XLAdresInfo_20220 adresInfo, string name)
        {
            int[] index = new int[2] { -1, -1 };
            if (name.Length > 50)
            {
                index[0] = name.Substring(0, 50).LastIndexOf(" ");
                if (index[0] > 0)
                {
                    index[1] = name.Substring(index[0] + 1).LastIndexOf(" ");
                }
                else
                    index[1] = name.Substring(50, 50).LastIndexOf(" ");
                if (index[0] > 0)
                {
                    adresInfo.Nazwa1 = name.Substring(0, index[0]);
                    adresInfo.Nazwa2 = name.Substring(index[0] + 1);
                    if (name.Substring(index[0] + 1).Length > 50)
                    {
                        adresInfo.Nazwa3 = name.Substring(index[0] + 51);
                    }
                }
                else
                {
                    adresInfo.Nazwa1 = name.Substring(0, 50);
                    adresInfo.Nazwa2 = name.Substring(51);
                    if (name.Length > 100)
                    {
                        adresInfo.Nazwa3 = name.Substring(index[0] + 51);
                    }

                }
            }
            else { adresInfo.Nazwa1 = name; }
        }
        public void SplitContractorName(XLKontrahentInfo_20220 kontrahentInfo, string name)
        {
            int[] index = new int[2] {-1,-1};
            if (name.Length > 50)
            {
                index[0] = name.Substring(0, 50).LastIndexOf(" ");
                if (index[0] > 0)
                {
                    index[1] = name.Substring(index[0] + 1).LastIndexOf(" ");
                }
                else
                    index[1] = name.Substring(50, 50).LastIndexOf(" ");
                if (index[0] > 0)
                {
                    kontrahentInfo.Nazwa1 = name.Substring(0, index[0]);
                    kontrahentInfo.Nazwa2 = name.Substring(index[0] + 1);
                    if (name.Substring(index[0] + 1).Length > 50)
                    {
                        kontrahentInfo.Nazwa3 = name.Substring(index[0] + 51);
                    }
                }
                else
                {
                    kontrahentInfo.Nazwa1 = name.Substring(0, 50);
                    kontrahentInfo.Nazwa2 = name.Substring(51);
                    if (name.Length > 100)
                    {
                        kontrahentInfo.Nazwa3 = name.Substring(index[0] + 51);
                    }

                }
            }
            else { kontrahentInfo.Nazwa1 = name; }
        }
        public int[] AddContractor_XL(int order_nr, TextWriter tw_log)
        {
            AttachThreadToClarion(true);
            //dodanie kontrahenta
            //todo
            //invoice comp - podm gospodarczy
            //
            string str = "";
            string kraj = ConnectToSQL(connection_string, "Select BL_Kraje.[Kod ISO] from ORD.BL_Kraje where BL_Kraje.Kraj = '" + cumulative_response.Orders[order_nr].InvoiceCountry + "'");
            string order_id = "";
            XLKontrahentInfo_20220 kontrahentInfo = new XLKontrahentInfo_20220
            {
                Wersja   = XLWersjaAPI,
                Akronim  = "Knt_" + (Int32.Parse(ConnectToSQL(connection_string, "Select Max(KntKarty.Knt_GIDNumer) + 1 from [cdn].[KntKarty]"))).ToString(),
                NipE     = Regex.Replace(cumulative_response.Orders[order_nr].InvoiceNip, @"[^\d]", ""),
                Ulica    = cumulative_response.Orders[order_nr].InvoiceAddress,
                Miasto   = cumulative_response.Orders[order_nr].InvoiceCity,
                KodP     = cumulative_response.Orders[order_nr].InvoicePostcode,
                Kraj     = kraj,
                EMail    = cumulative_response.Orders[order_nr].Email,
                Telefon1 = cumulative_response.Orders[order_nr].Phone,
                GrupaSciezka = @"E-SKLEP"
            };
            if (kontrahentInfo.NipE.Length > 1 && kontrahentInfo.NipE.Length >= 10)
            {
                kontrahentInfo.NipE = kontrahentInfo.NipE.PadRight(10, '0');
                kontrahentInfo.Status = 1;
                kontrahentInfo.PlatnikVat = 1;
            }
            if (cumulative_response.Orders[order_nr].InvoiceFullname=="")
            {
                SplitContractorName(kontrahentInfo, cumulative_response.Orders[order_nr].InvoiceCompany);
            }
            else
            {
                SplitContractorName(kontrahentInfo, cumulative_response.Orders[order_nr].InvoiceFullname);
                kontrahentInfo.Nazwa3 += " " + cumulative_response.Orders[order_nr].InvoiceCompany;
            }

            if (cumulative_response.Orders[order_nr].InvoiceAddress.Length>50)
            {
                int last_index = cumulative_response.Orders[order_nr].InvoiceAddress.Substring(0,50).LastIndexOf(" ");
                if (last_index > 0)
                {
                    kontrahentInfo.Ulica = cumulative_response.Orders[order_nr].InvoiceAddress.Substring(0, last_index);
                    kontrahentInfo.Adres = cumulative_response.Orders[order_nr].InvoiceAddress.Substring(last_index + 1);
                }
                else kontrahentInfo.Adres = cumulative_response.Orders[order_nr].InvoiceAddress.Substring(51);

            }
            try
            {
                str = cdn_api.cdn_api.XLNowyKontrahent(sesja, ref XLkontrahent_id, kontrahentInfo).ToString();
            }
            catch (Exception e)
            {
                tw_log.WriteLine(DateTime.Now + " Wystąpił błąd " + str + " dla kontrahenta o id " + kontrahentInfo.Akronim + " wyjątek : " + e);
            }
            finally
            {
                if (str != "0")
                {
                    tw_log.WriteLine("Wystąpił błąd API o numerze " + str + ", dla kontrahentInfo nr " + kontrahentInfo.GIDNumer.ToString() + " z nip " + kontrahentInfo.NipE);
                }
            }
            XLAdresInfo_20220 xLAdresInfo = new XLAdresInfo_20220
            {
                Wersja = XLWersjaAPI,
                Wysylkowy = 1,
                KntTyp = kontrahentInfo.GIDTyp,
                KntFirma = kontrahentInfo.GIDFirma,
                KntNumer = kontrahentInfo.GIDNumer,
                KntLp = kontrahentInfo.GIDLp,
                Miasto = cumulative_response.Orders[order_nr].DeliveryCity,
                KodP = cumulative_response.Orders[order_nr].DeliveryPostcode,
                Ulica = cumulative_response.Orders[order_nr].DeliveryAddress,
                Kraj = kraj,
                //Nazwa1 = cumulative_response.Orders[order_nr].DeliveryFullname,
                //Nazwa2 = cumulative_response.Orders[order_nr].DeliveryCompany,
            };
            if (cumulative_response.Orders[order_nr].DeliveryFullname != "")
            {
                SplitDeliveryName(xLAdresInfo, cumulative_response.Orders[order_nr].DeliveryFullname);
            } else
            {
                SplitDeliveryName(xLAdresInfo, cumulative_response.Orders[order_nr].DeliveryCompany);
            }
            if (kontrahentInfo.Status == 1)
            { 
                if (cumulative_response.Orders[order_nr].DeliveryFullname != "") 
                     xLAdresInfo.Nazwa3 += cumulative_response.Orders[order_nr].DeliveryCompany;
                else xLAdresInfo.Nazwa3 += cumulative_response.Orders[order_nr].InvoiceCompany;
            }
            if (cumulative_response.Orders[order_nr].DeliveryAddress.Length > 50)
            {
                int last_index = cumulative_response.Orders[order_nr].DeliveryAddress.Substring(0, 50).LastIndexOf(" ");
                if (last_index > 0)
                {
                    kontrahentInfo.Ulica = cumulative_response.Orders[order_nr].DeliveryAddress.Substring(0, last_index);
                    kontrahentInfo.Adres = cumulative_response.Orders[order_nr].DeliveryAddress.Substring(last_index + 1);
                }
                else kontrahentInfo.Adres = cumulative_response.Orders[order_nr].DeliveryAddress.Substring(51);

            }
            try
                {
                    str = cdn_api.cdn_api.XLNowyAdres(sesja, ref XLadres_id, xLAdresInfo).ToString();
                }
                catch (Exception e)
            {

                str += "Wystąpił błąd " + str + " dla adresu kontrahenta o id " + kontrahentInfo.Akronim + " wyjątek : " + e;
                throw new Exception("Wystąpił błąd dodawania adresu kontrahenta nr " + str + " dla kontrahenta o id " + xlDokumentZamInfo.GIDNumer + " wyjątek : " + e);
            }
            finally
            {
                if (str != "0")
                {
                    Console.WriteLine(DateTime.Now + " Wystąpił błąd API o numerze " + str + ", dla dodawania adresu dla kontrahenta o nr " + kontrahentInfo.GIDNumer.ToString());
                    tw_log.WriteLine(DateTime.Now + " Wystąpił błąd API o numerze " + str + ", dla dodawania adresu dla kontrahenta o nr " + kontrahentInfo.GIDNumer.ToString());
                }
            }
            return new int[] { kontrahentInfo.GIDTyp, kontrahentInfo.GIDFirma, kontrahentInfo.GIDNumer, kontrahentInfo.GIDLp };
        }

        public string AddZSOrder_XL(int order_nr, int[] response,string order_id, TextWriter tw_log)
        {
            string str = "";
            try
            {
                AttachThreadToClarion(true);
                
                string akronim = "BL_";
                int tryb = 0;
                string cecha = ConnectToSQL(connection_string, "Select BL_Zrodla.WartoscXL" +
                                                               " from ORD.BL_Zrodla where " +
                                                               "BL_Zrodla.WartoscBL =" + Int32.Parse(cumulative_response.Orders[order_nr].OrderSourceId));
                string numer = "";
                numer = ConnectToSQL(connection_string, "Select BL_Zrodla.SerieZS" +
                       " from ORD.BL_Zrodla where " +
                       "BL_Zrodla.WartoscBL =" + Int32.Parse(cumulative_response.Orders[order_nr].OrderSourceId));

                string msg = "";
                string kraj = "";
                int formaPlatnosci = 60;
                if (!String.IsNullOrEmpty(msg = ConnectToSQL(connection_string, "Select [ORD].[BL_RodzajTransakcji].[FormaPl]" +
                                                              " from [ORD].[BL_RodzajTransakcji] where " +
                                                              "[BL_RodzajTransakcji].[MetodaPlatnosciBL] = '" + cumulative_response.Orders[order_nr].PaymentMethod + "'")))
                {
                    if (Int32.TryParse(msg, out formaPlatnosci) != true)
                    {
                        Console.WriteLine(msg);

                        Console.WriteLine("Parse formy płatności nie powiodło się.");
                    }

                }
                if (cumulative_response.Orders[order_nr].WantInvoice == "1")
                {
                    if (cumulative_response.Orders[order_nr].InvoiceFullname != "")
                        akronim = cumulative_response.Orders[order_nr].InvoiceFullname;
                    else if (cumulative_response.Orders[order_nr].InvoiceCompany != "")
                        akronim = cumulative_response.Orders[order_nr].InvoiceCompany;
                }
                else if (cumulative_response.Orders[order_nr].DeliveryFullname != "")
                    akronim = cumulative_response.Orders[order_nr].DeliveryFullname;
                else akronim = cumulative_response.Orders[order_nr].Email;
                kraj = ConnectToSQL(connection_string, "Select BL_Kraje.[Kod ISO] from ORD.BL_Kraje where BL_Kraje.Kraj = '" + cumulative_response.Orders[order_nr].DeliveryCountry + "'");
                //Danie jutrzejszej daty dokumentu jeśli jest godz. 19-24
                DateTime docDate = DateTime.Now;
                if ((docDate.Hour == 18 && docDate.Minute >= 50) || docDate.Hour > 18)
                {
                    docDate = docDate.Date.AddDays(1).AddHours(6);
                }
                XLDokumentZamNagInfo_20220 xlDokumentZamInfo = new XLDokumentZamNagInfo_20220
                {
                    Wersja = XLWersjaAPI,
                    Magazyn = "M" + numer,
                    ZamSeria = "S" + numer,
                    Akronim = akronim,
                    Typ = 6, //ZS
                    Opis = "Uwagi klienta : " + cumulative_response.Orders[order_nr].UserComments,// + "\n Uwagi sklepu : " + cumulative_response.Orders[order_nr].AdminComments,
                    Waluta = cumulative_response.Orders[order_nr].Currency,
                    FormaPl = formaPlatnosci,
                    KntTyp = response[0],
                    KntFirma = response[1],
                    KntNumer = response[2],
                    KntLp = response[3],
                    //Rok = DateTime.Now.Year,
                    //Miesiac = DateTime.Now.Month,
                    SposobDst = cumulative_response.Orders[order_nr].DeliveryMethod,
                    KodRodzajuTransportu = "3",
                    TerminPlatnosci = Int32.Parse(cumulative_response.Orders[order_nr].PaymentMethodCod),
                    DataPotwierdzeniaOferty = GetClarionDate(docDate),
                    DataWystawienia = GetClarionDate(docDate),
                    Kraj = kraj,
                    Cecha = cecha,
                    DokumentObcy = "ESKLEP " + cumulative_response.Orders[order_nr].OrderId,
                    FlagaNB = "B",
                    PotwierdzenieOferty = 1,
                    URL = cumulative_response.Orders[order_nr].OrderPage,
                    DomyslnyDokRealizacja = "2033",
                };
                if (cumulative_response.Orders[order_nr].AdminComments != "")
                    xlDokumentZamInfo.Opis += "\n Uwagi sklepu : " + cumulative_response.Orders[order_nr].AdminComments;
                if (double.TryParse(String.Concat(cumulative_response.Orders[order_nr].ExtraField2.Where(Char.IsDigit)), out double result))
                {
                    xlDokumentZamInfo.DataRealizacji =
                        Int32.Parse
                                (GetClarionDate
                                    (UnixTimeStampToDateTime
                                        (double.Parse(cumulative_response.Orders[order_nr].DateConfirmed) + (result * 24 * 3600))).ToString());
                }
                if (Int32.TryParse(
                    ConnectToSQL(connection_string, "Select BL_Zrodla.IdSklepuERP" +
                       " from ORD.BL_Zrodla where " +
                       "BL_Zrodla.WartoscBL =" + Int32.Parse(cumulative_response.Orders[order_nr].OrderSourceId)),
                    out int val))
                {
                    xlDokumentZamInfo.FrsID = val;
                }

                try
                {
                    str = cdn_api.cdn_api.XLNowyDokumentZam(sesja, ref XLdokument_id, xlDokumentZamInfo).ToString();
                }
                catch (Exception e)
                {
                    throw new Exception("Wystąpił błąd dodawania Zam " + str + " dla kontrahenta o id " + xlDokumentZamInfo.GIDNumer + " wyjątek : " + e);

                }
                if (str != "0")
                {

                    tw_log.WriteLine(DateTime.Now + " Wystąpił błąd dodawania Zam " + str + " dla kontrahenta o id " + xlDokumentZamInfo.GIDNumer);
                    return "-1";
                }
                Console.WriteLine($@"2. Dodanie ZS z order id {order_id}. API zwróciło kod: " + str + " a DokumentInfo to " + xlDokumentZamInfo.GIDNumer.ToString());
                tw_log.WriteLine(DateTime.Now + $@"2. Dodanie ZS z order id {order_id} API zwróciło kod: " + str + " a DokumentInfo to " + xlDokumentZamInfo.GIDNumer.ToString());
                if (cumulative_response.Orders[order_nr].PaymentMethodCod == "0")
                {


                    XLDokumentZamPlatInfo_20220 xlDokumentZamPlatnoscInfo = new XLDokumentZamPlatInfo_20220
                    {
                        Wersja = XLWersjaAPI,
                        Kwota = cumulative_response.Orders[order_nr].PaymentDone.ToString().Replace(".", ","),
                        Zaliczka = 1,
                    };
                    try
                    {
                        str = cdn_api.cdn_api.XLDodajPlatnoscZam(XLdokument_id, xlDokumentZamPlatnoscInfo).ToString();
                    }
                    catch (Exception e)
                    {
                        throw new Exception("Wystąpił błąd dodawania zaliczki zam " + str + " dla kontrahenta o id " + xlDokumentZamInfo.GIDNumer + " wyjątek : " + e);
                    }
                    Console.WriteLine("2A. Dodanie Zaliczki ZS. o kwocie " + xlDokumentZamPlatnoscInfo.Kwota + " API zwróciło kod: " + str + " a DokumentInfo to " + xlDokumentZamInfo.GIDNumer.ToString());
                }
                //Dodanie opłaty za transport, jesli COD != 0
                //

                Console.WriteLine("3. Ilosc elementow to " + cumulative_response.Orders[order_nr].Products.Length);
                int[] gid;
                for (int product_nr = 0; product_nr < cumulative_response.Orders[order_nr].Products.Length; product_nr++)
                {
                    gid = program.AddIndex_XL(order_nr, product_nr, product_nr.ToString(), xlDokumentZamInfo.GIDNumer, numer, tw_log);
                    if (gid[2] < 0)
                    {
                        return "-1";
                    }
                    Console.WriteLine("Dodano element nr {0}, API AddElement zwrocilo {1}",
                            product_nr + 1,
                            program.AddElement_XL(XLdokument_id, order_nr, product_nr, gid, numer, tw_log));

                }
                if (cumulative_response.Orders[order_nr].DeliveryPrice > 0)
                {
                    AddTransportPayment_XL(XLdokument_id, order_nr, tw_log);
                    Console.WriteLine("5. Opłacono :" + cumulative_response.Orders[order_nr].PaymentDone);
                    tw_log.WriteLine(DateTime.Now + " 5. Opłacono :" + cumulative_response.Orders[order_nr].PaymentDone);
                }
                //AddPayment_XL(order_nr);

                XLZamkniecieDokumentuZamInfo_20220 xLZamkniecieDokumentuInfo_20220 = new XLZamkniecieDokumentuZamInfo_20220
                {
                    Wersja = XLWersjaAPI,
                    TrybZamkniecia = 2,
                };
                //WycenaWystawianiaSamegoZS - Póki co OFF

                if (ConnectToSQL(connection_string, $" select top 1 CzyBufor from ord.BL_Zrodla where WartoscBL={cumulative_response.Orders[order_nr].OrderSourceId}") == "1")
                {
                    xLZamkniecieDokumentuInfo_20220.TrybZamkniecia = 0;
                }

                if (ConnectToSQL(connection_string, "with t as (select name_for_customer from ord.BL_StatusID where id='" + cumulative_response.Orders[order_nr].OrderStatusId + "') " +
                                                    "select case when exists(select * from t where t.name_for_customer Like '%anulowane%') then '1' else '0' end as answer from t")
                == "1")
                {
                    xLZamkniecieDokumentuInfo_20220.TrybZamkniecia = 4;
                }
                cdn_api.cdn_api.XLZamknijDokumentZam(XLdokument_id, xLZamkniecieDokumentuInfo_20220);
                Console.WriteLine(ConnectToSQL(connection_string, "update cdn.ZamNag set ZaN_FrSID=" + xlDokumentZamInfo.FrsID + " where ZaN_DokumentObcy='ESKLEP " + cumulative_response.Orders[order_nr].OrderId + "'") + "Gdzie FRSID to " + xlDokumentZamInfo.FrsID);

                string[] WartoscAtrybutu = new string[13] {"AUTOMAT",
                                        ConnectToSQL(connection_string,"declare @prd_name varchar(512), @prd_id int set @prd_id = "+cumulative_response.Orders[order_nr].OrderStatusId+" set @prd_name = '"+cumulative_response.Orders[order_nr].OrderStatusId+"'+'#NieznanyStatus'"+
                                        "begin if exists(Select 1 from ord.BL_StatusID where id = @prd_id)" +
                                        "begin set @prd_name = (Select name from ord.BL_StatusID where id = @prd_id) end " +
                                        "select @prd_name end"),
                                        cumulative_response.Orders[order_nr].WantInvoice == "1" ? "TAK" : "NIE",
                                        cumulative_response.Orders[order_nr].OrderId,
                                        cumulative_response.Orders[order_nr].ShopOrderId,
                                        cecha,
                                        cumulative_response.Orders[order_nr].OrderPage,
                                        cumulative_response.Orders[order_nr].UserLogin,
                                        cumulative_response.Orders[order_nr].DeliveryMethod,
                                        cumulative_response.Orders[order_nr].PaymentMethod,
                                        cumulative_response.Orders[order_nr].PaymentMethodCod == "1" ? "Płatność Przy odbiorze" : (cumulative_response.Orders[order_nr].PaymentDone > 0 ? "Opłacone" : "Nieopłacone"),
                                        "-",
                                        cumulative_response.Orders[order_nr].DeliveryPackageModule
                                                      };
                Console.WriteLine(WartoscAtrybutu[1]);
                tw_log.WriteLine("Status zamówienia w atrybucie: " + cumulative_response.Orders[order_nr].OrderStatusId);

                var czesciowaRealizacja = xLZamkniecieDokumentuInfo_20220.TrybZamkniecia == 0 ? "TAK" : "NIE";
                Console.WriteLine("Następuje dodawanie atrybutów");
                tw_log.WriteLine("Dodawanie atrybutów. " + program.AddAtribute(order_nr.ToString(), WartoscAtrybutu, czesciowaRealizacja));
                if (cumulative_response.Orders[order_nr].PaymentDone != 0 &&
                    (cumulative_response.Orders[order_nr].InvoiceNip.Length == 10 || cumulative_response.Orders[order_nr].InvoiceNip.Length < 2) &&
                    xLZamkniecieDokumentuInfo_20220.TrybZamkniecia != 4 &&
                    xLZamkniecieDokumentuInfo_20220.TrybZamkniecia != 0)

                {
                    //do przerobienia. WartoscAtrybutu[1] = "7 faktury do wystawienia"
                    //WartoscAtrybutu[1] = "7 faktury do wystawienia"; 
                    if (ConnectToSQL(connection_string, "select top 1 BL_Zrodla.CzyFSL from ord.BL_Zrodla where WartoscBL='" + cumulative_response.Orders[order_nr].OrderSourceId + "'") == "1")
                    {
                        tw_log.WriteLine($@"Dodawanie faktury FSL do zamówienia ESKLEP {cumulative_response.Orders[order_nr].OrderId} dla dokumentu {xlDokumentZamInfo.GIDNumer}");
                        program.AddFSL_XL(order_nr, response, 1824, new int[4] { xlDokumentZamInfo.GIDTyp, xlDokumentZamInfo.GIDTyp, xlDokumentZamInfo.GIDNumer, xlDokumentZamInfo.GIDLp }, kraj, tw_log);
                        Console.WriteLine(ConnectToSQL(connection_string,
                        "update cdn.TraNag set TrN_FrsID=(select ZaN_FrsID from cdn.ZamNag where ZaN_DokumentObcy='ESKLEP " + cumulative_response.Orders[order_nr].OrderId + "') where TrN_ZaNNumer=(select ZaN_GIDNumer from cdn.ZamNag where ZaN_DokumentObcy='ESKLEP " + cumulative_response.Orders[order_nr].OrderId + "')"
                        ) + " Zaktualizowano FrsID dla FSL/FS");

                    }
                }
                else tw_log.WriteLine($@"Nie dodaje faktury do ESKLEP {cumulative_response.Orders[order_nr].OrderId} dla dokumentu {xlDokumentZamInfo.GIDNumer}");
            }
            catch(Exception en )
            {
                Console.WriteLine(" Dla kontrahenta o id " + xlDokumentZamInfo.GIDNumer + " wyjątek : " + en);
            }

            return str;
        }
        
        public string AddElement_XL(int IDDokumentuZam, int order_nr, int product_nr, int[] index, string numer, TextWriter tw_log)
        {
            string str = "";
            XLDokumentZamElemInfo_20220 xlDokumentZamInfo = new XLDokumentZamElemInfo_20220
            {
                Wersja = XLWersjaAPI,
                Ilosc = cumulative_response.Orders[order_nr].Products[product_nr].Quantity,
                Nazwa = cumulative_response.Orders[order_nr].Products[product_nr].Name,
                StawkaVAT = cumulative_response.Orders[order_nr].Products[product_nr].TaxRate,
                Magazyn = "M" + numer,//cumulative_response.Orders[order_nr].Products[product_nr].StorageId,
                //Wartosc = (cumulative_response.Orders[order_nr].Products[product_nr].PriceBrutto).ToString(),
                TwrTyp = index[0],
                TwrFirma = index[1],
                TwrNumer = index[2],
                TwrLp = index[3],
                //MagNumer = cumulative_response.Orders[order_nr].Products[product_nr].StorageId, - nie jestem pewien tego rozwiązania.
                Cecha = cumulative_response.Orders[order_nr].Products[product_nr].Attributes
            };
            switch (cumulative_response.Orders[order_nr].Products[product_nr].TaxRate)
            {
                case "23":
                    xlDokumentZamInfo.GrupaVAT = "A";
                    break;
                case "8":
                    xlDokumentZamInfo.GrupaVAT = "B";
                    break;
                case "3":
                    xlDokumentZamInfo.GrupaVAT = "F";
                    break;
                case "0":
                    xlDokumentZamInfo.GrupaVAT = "D";
                    break;
                default:
                    xlDokumentZamInfo.GrupaVAT = "A";
                    break;
            }
            try
            {
                str = cdn_api.cdn_api.XLDodajPozycjeZam(IDDokumentuZam, xlDokumentZamInfo).ToString();
            }
            catch (Exception e)
            {
                throw new Exception("Błąd nr " + e.ToString() + " kod API: " + str);
            }
            finally
            {
                dokInfo = new int[4] { xlDokumentZamInfo.GIDFirma, xlDokumentZamInfo.GIDLp, xlDokumentZamInfo.GIDNumer, xlDokumentZamInfo.GIDTyp };
                tw_log.WriteLine(DateTime.Now + " 7. API po dodaniu ZamInfo zwrocilo : " + str);
                Console.WriteLine("7. API po dodaniu ZamInfo zwrocilo : " + str);
            }
            return str;
        }
        //dodawanie towaru z zamowienia do katalogu
        public int[] AddIndex_XL(int order_nr, int product_nr, string index_nr, int order_gid, string numer, TextWriter tw)
        {
            //System.Threading.Thread.Sleep(1200);
            //AttachThreadToClarion(true);
            Random rand = new Random();
            int return_value = -1;
            string nazwa = "";
            //Wydobywanie EAN
            if (cumulative_response.Orders[order_nr].Products[product_nr].Name.LastIndexOf("[EAN") != -1)
            { 
                nazwa = (cumulative_response.Orders[order_nr].Products[product_nr].Name).Substring(0, cumulative_response.Orders[order_nr].Products[product_nr].Name.LastIndexOf("[EAN"));
                if (nazwa.Last() == ' ')
                {
                    nazwa = nazwa.Substring(0, nazwa.Length - 1);
                }
            }
            else
                nazwa = (cumulative_response.Orders[order_nr].Products[product_nr].Name);
            
            //Nazwa producenta, ostatnie słowo
            Console.WriteLine(nazwa);
           // Console.ReadKey();
            Random rnd = new Random();
            //Na podstawie ostatniego zostanie wyznaczony producent, jeśli takowy istnieje w tabeli SQL. Substring, lastindexof ze spacji -> czy istnieje -> producent
            //Domyślnie ma to być Atrybut producent
            //Atrybut ID PROD na podsatwie ID Producenta
            XLTowarInfo_20220 xLTowarInfo = new XLTowarInfo_20220
            {
                Wersja = XLWersjaAPI,
                Typ = 1,
                Kod = cumulative_response.Orders[order_nr].OrderId + "_" + (product_nr + 1).ToString().PadLeft(2, '0'),
                StawkaPod = cumulative_response.Orders[order_nr].Products[product_nr].TaxRate,
                StawkaPodSpr = cumulative_response.Orders[order_nr].Products[product_nr].TaxRate,
                Nazwa = @nazwa,
                Magazyn = "M"+numer,
                EAN = "2010000000001",
                Jm = "szt",
                TwrGrupa = "M"
            };
            switch (cumulative_response.Orders[order_nr].Products[product_nr].TaxRate)
            {
                case "23":
                    xLTowarInfo.GrupaPod = "A";
                    xLTowarInfo.GrupaPodSpr = "A";
                    break;
                case "8":
                    xLTowarInfo.GrupaPod = "B";
                    xLTowarInfo.GrupaPodSpr = "B";
                    break;
                case "3":
                    xLTowarInfo.GrupaPod = "F";
                    xLTowarInfo.GrupaPodSpr = "F";
                    break;
                case "0":
                    xLTowarInfo.GrupaPod = "D";
                    xLTowarInfo.GrupaPodSpr = "D";
                    break;
                default:
                    xLTowarInfo.GrupaPod = "A";
                    xLTowarInfo.GrupaPodSpr = "A";
                    break;
            }
            /*
            string qr_name = ConnectToSQL(connection_string, "declare @a varchar(255) = '" + nazwa + "'" +
                @"select top 1 replace(cdn.TwrGrupaPelnaNazwa(AtH_ObiNumer), '/', '\|')" +
                @"from cdn.AtrybutyKlasy as k left join cdn.AtrybutyHist as h on k.AtK_ID = h.AtH_AtkId " +
                @"where AtK_Nazwa='Słowo kluczowe grupa' and @a like '%' + AtH_Wartosc + '%'");
            if (qr_name !="")
            {
                if (qr_name == "Krzesło") qr_name = @"M\|Krzesła";
                xLTowarInfo.TwrGrupa = qr_name;
            }
            */
            if (!String.IsNullOrWhiteSpace(xLTowarInfo.Kod))
            {
                string qr_name = "declare @prod_name varchar(512), @twr_kod varchar(512) set @prod_name ='" + cumulative_response.Orders[order_nr].AdminComments + "' set @twr_kod ='" + xLTowarInfo.Kod + "' " +
                "begin if exists(Select 1 from cdn.KntKarty where Knt_Akronim LIKE '%' + @prod_name + '%') " +
                "update cdn.TwrKarty set Twr_PrdNumer = (select top 1 Knt_GIDNumer from cdn.KntKarty where Knt_Akronim LIKE '%' + @prod_name + '%' and Knt_GIDTyp = 32), Twr_PrdFirma = 1034247, Twr_PrdTyp = 32 where Twr_Kod = @twr_kod end " +
                "begin if exists(Select 1 from cdn.Slowniki where SLW_SLSId = 103 and SLW_Kategoria='Marki towarów' and SLW_WartoscS=@prod_name) " +
                "update cdn.TwrKarty set Twr_MarkaId = (select SLW_ID from cdn.Slowniki where SLW_SLSId = 103 and SLW_Kategoria = 'Marki towarów' and SLW_WartoscS = @prod_name) where Twr_Kod = @twr_kod end";
                Console.WriteLine(ConnectToSQL(connection_string, qr_name));
                
            }
        
            if (cumulative_response.Orders[order_nr].Products[product_nr].Ean != "")
            {
                xLTowarInfo.EAN = cumulative_response.Orders[order_nr].Products[product_nr].Ean;
            }
            else
            {
                 string lastEANQuery = "SELECT TOP (1) Twr_Ean FROM [CDN].[TwrKarty] WHERE Twr_Ean like '2000%' ORDER BY Twr_Ean desc"; //stara wersja '201000000%'
               // tw.WriteLine("lastEAN 0-12: "+ConnectToSQL(connection_string, lastEANQuery).Substring(0, 12));
                Int64 nextEAN = Convert.ToInt64(ConnectToSQL(connection_string, lastEANQuery).Substring(0, 12))+1;
                //tw.WriteLine("nextEAN: " + nextEAN);
                //short sEAN = Convert.ToInt16(nextEAN);
                string sEAN = Convert.ToString(nextEAN);
                string modulo = WyliczSume(sEAN);
                //tw.WriteLine("sEAN: " + sEAN+" + suma: "+modulo);
                xLTowarInfo.EAN = sEAN + modulo;

                //xLTowarInfo.EAN = Convert.ToString(nextEAN).TrimEnd();
               // XLEANInfo_20220 ean = new XLEANInfo_20220();
                //ean.Wersja = XLWersjaAPI;
                //ean.Zakres = 2;
                //int return_val = cdn_api.cdn_api.XLGenerujEAN(ean);
                //xLTowarInfo.EAN = ean.EAN;
                tw.WriteLine("EAN dla nowego towaru: " + xLTowarInfo.EAN);
            }
            
           
            try
            {
                //tw.WriteLine("EAN dla nowego towaru przed dodaniem: " + xLTowarInfo.EAN);
                return_value = cdn_api.cdn_api.XLNowyTowar(sesja, ref XLtowar_id, xLTowarInfo);
                //tw.WriteLine("EAN dla nowego towaru po dodaniu: " + xLTowarInfo.EAN);
                if (return_value==83)
                {
                    string[] dostawcaArray = ConnectToSQL(connection_string, "Select CONCAT(Twr_GIDTyp, ';',Twr_GIDFirma, ';', Twr_GIDNumer, ';', Twr_GIDLp) from cdn.TwrKarty where cdn.TwrKarty.Twr_Kod='" + cumulative_response.Orders[order_nr].OrderId + "_" + (product_nr + 1).ToString().PadLeft(2, '0') + "'").Split(';');
                    int[] res = new int[4];
                    for (int i = 0; i < dostawcaArray.Length; i++)
                    {
                        int tmp;
                        if (int.TryParse(dostawcaArray[i], out tmp))
                             res[i]= tmp;

                    }
                    xLTowarInfo.GIDTyp = res[0];
                    xLTowarInfo.GIDFirma = res[1];
                    xLTowarInfo.GIDNumer = res[2];
                    xLTowarInfo.GIDLp = res[3];
                    return_value = 0;
                    tw.WriteLine(DateTime.Now + "Pobrano istniejący towar o GIDTyp/Firma/Nr/Lp: " + xLTowarInfo.GIDTyp + @"/" + xLTowarInfo.GIDFirma + @"/" + xLTowarInfo.GIDNumer +@"/"+xLTowarInfo.GIDLp+". Kod API: " + return_value);
                    Console.WriteLine(DateTime.Now + "Pobrano istniejący towar o GIDTyp/Firma/Nr/Lp: " + xLTowarInfo.GIDTyp + @"/" + xLTowarInfo.GIDFirma + @"/" + xLTowarInfo.GIDNumer + @"/" + xLTowarInfo.GIDLp + ". Kod API: " + return_value);
                    //return new int[4] { xLTowarInfo.GIDTyp, xLTowarInfo.GIDFirma, xLTowarInfo.GIDNumer, xLTowarInfo.GIDLp };
                }
                else tw.WriteLine(DateTime.Now  + "Dodano towar o gid " + xLTowarInfo.GIDNumer + ". EAN: "+xLTowarInfo.EAN+". Kod API: " + return_value);

            }
            catch (Exception e)
            {
                tw.WriteLine("Dodano towar o gid " + xLTowarInfo.GIDNumer + ". Kod API: " + return_value + "ORAZ BŁĄD " + e);
                System.Environment.Exit(1);
                throw new System.Exception("Wystapil wyjatek przy dodawaniu towaru. Kod API: "+ return_value + ", Wyjątek: "+e);
                
            }
                Console.WriteLine("Dodany towar : {0}", nazwa);
                Console.WriteLine("6. Indeks"+nazwa+" dodany do magazynu o gid " + xLTowarInfo.GIDNumer +" Kod API: " + return_value);
                string msg = ConnectToSQL(connection_string, "update cdn.TwrCeny set TwC_Wartosc='"+cumulative_response.Orders[order_nr].Products[product_nr].PriceBrutto.ToString().Replace(",",".")+"', TwC_Marza='0.00' where TwC_TwrNumer = (select Twr_GIDNumer from cdn.TwrKarty where Twr_Kod = '"+xLTowarInfo.Kod+"') and TwC_TwrLp=1");
                if (msg !="")
                {
                    tw.WriteLine(DateTime.Now + " Wystapil błąd SQL przy dodawaniu ceny do zamowienia o gid"+xLTowarInfo.GIDNumer+". error: " + msg);
                }
                XLZamkniecieTowaruInfo_20220 xLZamkniecieTowaruInfo_20220 = new XLZamkniecieTowaruInfo_20220
                {
                    Wersja = XLWersjaAPI
                };
                cdn_api.cdn_api.XLZamknijTowar(XLtowar_id, xLZamkniecieTowaruInfo_20220);

            //Dodawanie producenta do towaru
            string produc = ConnectToSQL(connection_string,
            "select Knt_Akronim from cdn.KntKarty where Knt_GIDNumer = (select top 1 AtH_ObiNumer from cdn.Atrybuty left join cdn.AtrybutyHist on Atr_AtkId = AtH_AtkId where AtH_Wartosc = '" + SeparateProduct(cumulative_response.Orders[order_nr].AdminComments, product_nr) + "')");
            tw.WriteLine("Wydobyty akronim z nazwy producenta to " + produc);
                    if (produc != "")
                        {
                        Console.WriteLine("wydobyty Akronim producenta&dostawcy (jesli jakas)" + produc);
                        string query = "declare @prod_name varchar(512) set @prod_name = '"+produc+"'"+
                        "begin update cdn.TwrKarty set Twr_PrdNumer = " +
                        "(select distinct top 1 Knt_GIDNumer from cdn.KntKarty where Knt_Akronim like '%' + @prod_name + '%' order by Knt_GIDNumer asc), " +
                        "Twr_PrdLp =(select distinct Knt_GIDLp from cdn.KntKarty where Knt_Akronim like '%' + @prod_name + '%'), " +
                        "Twr_PrdFirma = (select distinct Knt_GIDFirma from cdn.KntKarty where Knt_Akronim like '%' + @prod_name + '%'), " +
                        "Twr_PrdTyp = (select distinct Knt_GIDTyp from cdn.KntKarty where Knt_Akronim like '%' + @prod_name + '%') " +
                        "where Twr_Kod = '"+ xLTowarInfo.Kod+ "' end";

                        Console.WriteLine(ConnectToSQL(connection_string, query));
                string[] dostawcaArray = ConnectToSQL(connection_string, "select  CONCAT(Knt_Nazwa1, ';',Knt_GIDNumer) from cdn.KntKarty where Knt_Akronim='"+produc+"'").Split(';');
                for (int i = 0; i < dostawcaArray.Length; i++)
                {
                    Console.WriteLine(dostawcaArray[i]);
                }
                
                try
                {
                    XLTwDInfo_20220 xLdostawca = new XLTwDInfo_20220
                    {
                        Wersja = XLWersjaAPI,
                        TwrKod = xLTowarInfo.Kod,
                        TwrNumer = xLTowarInfo.GIDNumer,
                        TwrTyp = xLTowarInfo.Typ,
                        KntAkronim = produc,
                        Waluta = "PLN",
                        KntTyp = 32,
                        NazwaKnt = dostawcaArray[0],
                        KntNumer = Int32.Parse(dostawcaArray[1]),
                        ProponowanyDst = 1,
                    };
                    return_value=cdn_api.cdn_api.XLDodajDostawce(sesja, xLdostawca);
                    Console.WriteLine("Dodanie dostawcy do towaru zwróciło kod API " + return_value);
                    tw.WriteLine("Dodanie dostawcy do towaru zwróciło kod API " + return_value);
                    Console.WriteLine(ConnectToSQL(connection_string, "update t " +
                        "set Twr_DstDomyslny = 1 " +
                        "from cdn.TwrKarty as t " +
                        "where Twr_GIDTyp = 16 AND Twr_GIDNumer = " + xLTowarInfo.GIDNumer));
                }
                catch (Exception)
                {
                    Console.WriteLine("błąd przy dodawaniu dostawcy do towaru!");
                }
               
                        }

            return  new int[4] {xLTowarInfo.GIDTyp, xLTowarInfo.GIDFirma, xLTowarInfo.GIDNumer, xLTowarInfo.GIDLp};
        }

        private string SeparateProduct(string adminComments, int product_nr)
        {
            string producent = "";
            string[] stringSeparators = new string[] { "#" };
            string[] firstNames = adminComments.Split(stringSeparators, StringSplitOptions.None);
            if (firstNames.Length > product_nr)
            {
                try
                {
                    return producent = firstNames[product_nr + 1]; 
                }
                catch (Exception e)
                {

                    return "";
                }
            }
            return "";
        }

        public int AddPayment_XL(int order_nr)
        {
            AttachThreadToClarion(true);
            int return_value = -1;
            double kwota = 0;
            for (int i = 0; i < cumulative_response.Orders[order_nr].Products.Length; i++)
            {
                kwota += cumulative_response.Orders[order_nr].Products[i].PriceBrutto;
            }
            if (cumulative_response.Orders[order_nr].DeliveryPrice > 0)
            {
                kwota += cumulative_response.Orders[order_nr].DeliveryPrice;
            }
            
            XLDokumentZamPlatInfo_20220 xLPlatnoscInfo = new XLDokumentZamPlatInfo_20220
            {
                Wersja = XLWersjaAPI,
                Kwota = cumulative_response.Orders[order_nr].PaymentDone.ToString(),
            };
            try
            {
                return_value = cdn_api.cdn_api.XLDodajPlatnoscZam(XLdokument_id, xLPlatnoscInfo);
            }
            catch (Exception e)
            {
                throw new System.Exception("Wystapil wyjatek przy dodawaniu płatności. Kod API: " + return_value + ", Wyjątek: " + e);
            }
            finally
            {
                if(return_value!=0)
                {
                    throw new System.Exception("Wystąpił błąd przy dodawaniu płatności, Kod API:" + return_value);
                }
            }
            
            return return_value;
        }

        public string AddTransportPayment_XL(int IDDokumentuZam, int order_nr, TextWriter tw_log)
        {
            AttachThreadToClarion(true);
            string str = "";
            XLDokumentZamElemInfo_20220 xlDokumentZamInfo = new XLDokumentZamElemInfo_20220
            {
                Wersja =    XLWersjaAPI,
                Ilosc =     "1",
                Wartosc =   (cumulative_response.Orders[order_nr].DeliveryPrice).ToString(),
                Nazwa =     "Usługi transportowe",
                //StawkaVAT = "23",
                TwrNumer =  95,
                TwrLp =     0,
                TwrTyp =    4,
                Towar =     "US.TRANSP.",
                Magazyn =   "MAG"
                //MagNumer = cumulative_response.Orders[order_nr].Products[product_nr].StorageId, - nie jestem pewien tego rozwiązania.
            };

            try
            {
                str = cdn_api.cdn_api.XLDodajPozycjeZam(IDDokumentuZam, xlDokumentZamInfo).ToString();
            }
            catch (Exception e)
            {
                throw new Exception("Błąd dla oplaty transportowej nr " + e.ToString() + " kod API: " + str);
            }
            finally
            {
                if (str != "0")
                {
                    tw_log.WriteLine(DateTime.Now + " Błąd dla oplaty transportowej, kod API: " + str);
                    Console.WriteLine(DateTime.Now + " Błąd dla oplaty transportowej, kod API: " + str);
                }
                else
                {
                Console.WriteLine("4. API po dodaniu oplaty transportowej zwrocilo : " + str);
                }

            }
            return str;
        }
        public string AddAtribute(string order_nr, string[] WartoscAtrybutu, string czesciowaRealizacja)
        {
            string resp = "";
            ConnectToSQL(connection_string,"if not exists (select 1 from cdn.Atrybuty where Atr_ObiTyp = 960 AND Atr_ObiNumer = " + dokInfo[2].ToString() + " AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND Atr_AtkId = " + KlasyAtrybutu[0].ToString() + ") begin " +
            "insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) " +
                                "VALUES(" + dokInfo[3].ToString() + ", " +
                                dokInfo[0].ToString() + ", " +
                                dokInfo[2].ToString() + ", " +
                                "0, " +
                                "0," + "(Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='" + KlasyAtrybutuNazwa[1] + "')" + ",'" + WartoscAtrybutu[1] + "', 0, 0, 0, 0, 0, 0, 0,1); end");
            string ifnotexists = "if not exists (select 1 from cdn.Atrybuty where Atr_ObiTyp = 960 AND Atr_ObiNumer = " + dokInfo[2].ToString() + " AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND (Atr_AtkId = " + KlasyAtrybutu[1].ToString();
            for (int i = 2; i < WartoscAtrybutu.Length; i++)
            {
                ifnotexists += " OR Atr_AtkID = " + "(Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='" + KlasyAtrybutuNazwa[i] + "')";
            }
            ifnotexists += ")) begin ";
            string query = ifnotexists + "insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) " +
                                "VALUES(" + dokInfo[3].ToString() + ", " +
                                dokInfo[0].ToString() + ", " +
                                dokInfo[2].ToString() + ", " +
                                "0, " +
                                "0," + "(Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='" + KlasyAtrybutuNazwa[1] + "')" + ",'" + WartoscAtrybutu[1] + "', 0, 0, 0, 0, 0, 0, 0,1)";
            for (int i = 2; i < WartoscAtrybutu.Length; i++)
            {
                query += ",(" + dokInfo[3].ToString() + ", " +
                                dokInfo[0].ToString() + ", " +
                                dokInfo[2].ToString() + ", " +
                                "0, " +
                                "0," + "(Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa='" + KlasyAtrybutuNazwa[i] + "')" + ",'" + WartoscAtrybutu[i] + "', 0, 0, 0, 0, 0, 0, 0,1)";
            }
            query += "; end";
            resp =  ConnectToSQL(connection_string, query);
            if (czesciowaRealizacja == "TAK")
            {
            resp += ConnectToSQL(connection_string, $@"if not exists (select 1 from cdn.Atrybuty right join cdn.AtrybutyKlasy on Atk_Id = Atr_AtkId where Atr_ObiTyp = {dokInfo[3]} AND Atr_ObiNumer = {dokInfo[2]} AND Atr_ObiLp = 0 AND Atr_ObiSubLp = 0 AND (AtK_Nazwa = 'Czesciowa Realizacja')) 
            begin 
            insert into cdn.Atrybuty(Atr_ObiTyp, Atr_ObiFirma, Atr_ObiNumer, Atr_ObiLp, Atr_ObiSubLp, Atr_AtkId, Atr_Wartosc, Atr_AtrTyp, Atr_AtrFirma, Atr_AtrNumer, Atr_AtrLp, Atr_AtrSubLp, Atr_OptimaId, Atr_Grupujacy, Atr_Pozycja) 
            VALUES({dokInfo[3]},
                   {dokInfo[0]},
                   {dokInfo[2]},
                        0,
                        0, (Select AtK_ID from cdn.AtrybutyKlasy where AtK_Nazwa = 'Czesciowa Realizacja'), 'TAK', 0, 0, 0, 0, 0, 0, 0, 1) 
                    end; ");
            
            }
            return resp;
        }
        public int AddFSL_XL(int order_nr, int[] kontrahent, int typ, int[] zamowienie, string kraj, TextWriter tw_log)
        {
            int return_value = -1;
            int tryb = 0;
            string cecha = ConnectToSQL(connection_string, "Select BL_Zrodla.WartoscXL" +
                                               " from ORD.BL_Zrodla where " +
                                               "BL_Zrodla.WartoscBL =" + Int32.Parse(cumulative_response.Orders[order_nr].OrderSourceId));
            DateTime docDate = DateTime.Now;
            if ((docDate.Hour >= 18 && docDate.Minute >= 50) || docDate.Hour>18)
            {
                docDate = docDate.Date.AddDays(1).AddHours(6);
            }
            XLDokumentNagInfo_20220 xLDokumentNagInfo = new XLDokumentNagInfo_20220
            {
                Wersja = XLWersjaAPI,
                Typ = 1824, //typ,
                GIDNumer = 0,
                GIDFirma = 0,
                GIDLp = 0,
                GIDTyp = 0,
                KntTyp = kontrahent[0],
                KntFirma = kontrahent[1],
                KntNumer = kontrahent[2],
                KntLp = kontrahent[3],
                ZamTyp = zamowienie[0],
                ZamFirma = zamowienie[1],
                ZamNumer = zamowienie[2],
                ZamLp = zamowienie[3],
                Tryb = 2,
                Rozliczac = 1, // No nie wiem
                //Fiskalny = 0,
                Rejestr = "SDET",
                DeklaracjaVAT7 = 2,
                NB = "B",
                //Rok = UnixTimeStampToDateTime(double.Parse(cumulative_response.Orders[order_nr].DateAdd)).Year,
                //Miesiac = UnixTimeStampToDateTime(double.Parse(cumulative_response.Orders[order_nr].DateAdd)).Month,
                Data = GetClarionDate(docDate),
                DataSpr = GetClarionDate(docDate),
                SposobDst = cumulative_response.Orders[order_nr].DeliveryMethod,
                Kraj = kraj,
                Cecha = cecha,
                Opis ="Uwagi klienta : " + cumulative_response.Orders[order_nr].UserComments,// + "\n Uwagi sklepu : " + cumulative_response.Orders[order_nr].AdminComments,

            };

            try
            {
                
                return_value = cdn_api.cdn_api.XLNowyDokument(sesja, ref XLFaktura_id,  xLDokumentNagInfo);
            }
            catch (Exception e)
            {
                throw new System.Exception("Wystapil wyjatek przy dodawaniu faktury. Kod API: " + return_value + ", Wyjątek: " + e);
            }
            finally
            {
                Console.WriteLine("9A. API po dodaniu faktury zwrocilo : " + return_value);
            }
            if(return_value!=0)
            {
                tw_log.WriteLine("9A. Dodanie FSL. API zwróciło kod: " + return_value + " a DokumentInfo to " + zamowienie[2].ToString());
                //tryb = -1;
                //return return_value;
                xlDokumentZamInfo.Miesiac = DateTime.Now.Month;
                return_value = cdn_api.cdn_api.XLNowyDokument(sesja, ref XLFaktura_id, xLDokumentNagInfo);
                tw_log.WriteLine("Próba dodania FSL pod innym miesiacem " + return_value + " a DokumentInfo to " + zamowienie[2].ToString());
                if (return_value!=0)
                {
                    tw_log.WriteLine("FSL z innym miesiacem sie pod innym miesiacem " + return_value + " a DokumentInfo to " + zamowienie[2].ToString());
                    tryb = -1;
                    return return_value;
                } 
            }


            List<double> vaty = new List<double>();
            if (cumulative_response.Orders[order_nr].DeliveryPrice > 0)
                vaty.Add(23.0);
            for (int i = 0; i < cumulative_response.Orders[order_nr].Products.Length; i++)
            {
                if (!vaty.Contains(Convert.ToDouble(cumulative_response.Orders[order_nr].Products[i].TaxRate)))
                {

                    vaty.Add(Convert.ToDouble(cumulative_response.Orders[order_nr].Products[i].TaxRate));
                }
            }

            

            for (int v = 0; v < vaty.Count(); v++)
            {
                double wartosc = 0;
                double wartoscVat = 0;
                for (int i = 0; i < cumulative_response.Orders[order_nr].Products.Length; i++)
                {
                    if (vaty[v] == Convert.ToDouble(cumulative_response.Orders[order_nr].Products[i].TaxRate))
                    {
                        wartoscVat += (cumulative_response.Orders[order_nr].Products[i].PriceBrutto* Convert.ToDouble(cumulative_response.Orders[order_nr].Products[i].Quantity)) - (cumulative_response.Orders[order_nr].Products[i].PriceBrutto* Convert.ToDouble(cumulative_response.Orders[order_nr].Products[i].Quantity)) / ((vaty[v] / 100) + 1);
                        wartosc += cumulative_response.Orders[order_nr].Products[i].PriceBrutto * Convert.ToDouble(cumulative_response.Orders[order_nr].Products[i].Quantity);
        }
                }
                if (vaty[v]==23.0 && cumulative_response.Orders[order_nr].DeliveryPrice > 0)
                {
                    wartoscVat += cumulative_response.Orders[order_nr].DeliveryPrice - cumulative_response.Orders[order_nr].DeliveryPrice / ((vaty[v] / 100) + 1);
                    wartosc += cumulative_response.Orders[order_nr].DeliveryPrice;
                }

                int rozliczac = 0;

                

                return_value = -1;

                

                XLVatInfo_20220 xLVatInfo_20220 = new XLVatInfo_20220
                {


                    Wersja = XLWersjaAPI,
                    FlagaVat = 1,
                    Rozliczac = rozliczac,
                    FlagaNB = "B",
                    //IgnorujVat = 1,
                    StawkaPod = Convert.ToInt32(vaty[v] * 100),
                    Wartosc = wartosc.ToString(),
                    WartoscVat = Math.Round(wartoscVat, 2, MidpointRounding.AwayFromZero).ToString(),//Math.Round(cumulative_response.Orders[order_nr].PaymentDone-cumulative_response.Orders[order_nr].PaymentDone/1.23, 2,MidpointRounding.AwayFromZero).ToString(),
                    DeklRok = xLDokumentNagInfo.Rok, //((UnixTimeStampToDateTime
                      //          (double.Parse(cumulative_response.Orders[order_nr].DateAdd))).Year),
                    DeklMiesiac = xLDokumentNagInfo.Miesiac, //((UnixTimeStampToDateTime
                      //          (double.Parse(cumulative_response.Orders[order_nr].DateAdd))).Month),

                };
                try
                {
                    return_value = cdn_api.cdn_api.XLDodajVAT(XLFaktura_id, xLVatInfo_20220);
                }
                catch (Exception e)
                {
                    throw new System.Exception("Wystapil wyjatek przy dodawaniu VAT faktury. Kod API: " + return_value + ", Wyjątek: " + e);
                }
                finally
                {

                    if (return_value != 0)
                    {
                        tryb = -1;
                        tw_log.WriteLine("Wystapil wyjatek przy dodawaniu VAT faktury. Kod API: " + return_value + " a DokumentInfo to " + xlDokumentZamInfo.GIDNumer.ToString());
                    }
                    Console.WriteLine("9B. API po dodaniu VAT do faktury zwrocilo : " + return_value);
                    tw_log.WriteLine("3B. API po dodaniu VAT do faktury zwrocilo : " + return_value);
                }
            }
            vaty.Clear();

            /* return_value = -1;
             XLVatInfo_20220 xLVatInfo_20220 = new XLVatInfo_20220
             {
                 Wersja = XLWersjaAPI,
                 FlagaVat = 1,
                 Rozliczac = 1,
                 FlagaNB = "B",
                 //IgnorujVat = 1,
                 StawkaPod = Int32.Parse(cumulative_response.Orders[order_nr].Products[0].TaxRate)*100,
                 Wartosc = (cumulative_response.Orders[order_nr].PaymentDone).ToString(),
                 WartoscVat = Math.Round(cumulative_response.Orders[order_nr].PaymentDone-cumulative_response.Orders[order_nr].PaymentDone/1.23, 2,MidpointRounding.AwayFromZero).ToString(),
                 DeklRok = ((UnixTimeStampToDateTime
                             (double.Parse(cumulative_response.Orders[order_nr].DateAdd))).Year),
                 DeklMiesiac = ((UnixTimeStampToDateTime
                             (double.Parse(cumulative_response.Orders[order_nr].DateAdd))).Month),

             };
             try
             {
                 return_value = cdn_api.cdn_api.XLDodajVAT(XLFaktura_id, xLVatInfo_20220);
             }
             catch (Exception e)
             {
                 throw new System.Exception("Wystapil wyjatek przy dodawaniu VAT faktury. Kod API: " + return_value + ", Wyjątek: " + e);
             }
             finally
             {

             if(return_value!= 0)
             {
                 tryb = -1;
                 tw_log.WriteLine("Wystapil wyjatek przy dodawaniu VAT faktury. Kod API: " + return_value + " a DokumentInfo to " + xlDokumentZamInfo.GIDNumer.ToString());
             }
                 Console.WriteLine("9B. API po dodaniu VAT do faktury zwrocilo : " + return_value);
                 tw_log.WriteLine("3B. API po dodaniu VAT do faktury zwrocilo : " + return_value);
             }*/
            return_value = -1;
            XLZamkniecieDokumentuInfo_20220 xLZamkniecieDokumentuInfo = new XLZamkniecieDokumentuInfo_20220
            {
                Wersja = XLWersjaAPI,
                GidNumer = xLDokumentNagInfo.GIDNumer,
                Tryb = 0,
                Fiskalny= 0, 
                LimitKnt = 2,
            };
            try
            {
                return_value = cdn_api.cdn_api.XLZamknijDokument(XLFaktura_id, xLZamkniecieDokumentuInfo);
            }
            catch (Exception e)
            {
                throw new System.Exception("Wystapil wyjatek przy dodawaniu faktury. Kod API: " + return_value + ", Wyjątek: " + e);
            }
            finally
            {

                Console.WriteLine("9C. API po zamknieciu faktury zwrocilo : " + return_value);
                tw_log.WriteLine("3C. API po zamknieciu faktury zwrocilo : " + return_value);
            }
            //Console.WriteLine(ConnectToSQL(connection_string, $"update cdn.TraNag set TrN_Rozliczac=0 where TrN_KntNumer={kontrahent[2]} and TrN_GIDTyp=1824"));
            Console.WriteLine(ConnectToSQL(connection_string, $@"declare @formanr varchar(128), @formarodz varchar(128) 
            set @formanr = (select top 1 ZaN_FormaNr from cdn.ZamNag where ZaN_GIDNumer={zamowienie[2]} and ZaN_GIDTyp={zamowienie[0]}) 
            set @formarodz = (select top 1 ZaN_FormaNazwa from cdn.ZamNag where ZaN_GIDNumer={zamowienie[2]} and ZaN_GIDTyp={zamowienie[0]}) 
            begin tran t1; 
            update cdn.TraNag set TrN_FormaNazwa=@formarodz, TrN_FormaNr=@formanr where TrN_GIDNumer={xLDokumentNagInfo.GIDNumer} and TrN_GIDTyp=1824 
            update cdn.TraPlat set TrP_FormaNazwa=@formarodz, TrP_FormaNr=@formanr where TrP_GIDNumer={xLDokumentNagInfo.GIDNumer} and TrP_GIDTyp=1824 
            commit;"));

            string nettoQuery = "SELECT SUM(TrN_NettoP + TrN_NettoR) FROM [CDN].[TraNag] WHERE [TrN_GIDTyp] = " + xLDokumentNagInfo.GIDTyp + " and [TrN_GIDFirma] = " + xLDokumentNagInfo.GIDFirma + " and [TrN_GIDNumer] = " + xLDokumentNagInfo.GIDNumer;
            double netto = Convert.ToDouble(ConnectToSQL(connection_string, nettoQuery));
            string fiskalnyQuery = "SELECT TrN_Fiskalny FROM [CDN].[TraNag] WHERE [TrN_GIDTyp] = " + xLDokumentNagInfo.GIDTyp + " and [TrN_GIDFirma] = " + xLDokumentNagInfo.GIDFirma + " and [TrN_GIDNumer] = " + xLDokumentNagInfo.GIDNumer;
            int fiskalny = Convert.ToInt32(ConnectToSQL(connection_string, fiskalnyQuery));
            tw_log.WriteLine("Netto = " + netto+"; Fiskalny = "+fiskalny) ;

            if (fiskalny == 0 && netto != 0)
            {
                ConnectToSQL(connection_string, $@"update cdn.TraVat set TrV_Rozliczac = 1  where TrV_GIDNumer = { xLDokumentNagInfo.GIDNumer } and TrV_GIDTyp = 1824");
            }

            Console.WriteLine("GID FAKTURY TO : " + xLDokumentNagInfo.GIDNumer);
            string karNr =
            "declare @xlkod varchar(128), @rodzpl varchar(128), @esklep varchar(128) " +
                "set @esklep = ('ESKLEP " + cumulative_response.Orders[order_nr].OrderId + "') " +
                "set @rodzpl = (select MetodaPlatnosciXL from ord.BL_RodzajTransakcji where MetodaPlatnosciBL = '" + cumulative_response.Orders[order_nr].PaymentMethod + "') " +
                "set @xlkod =  (select KAR_GIDNumer from cdn.Rejestry where KAR_Seria = (SELECT XLKod from ord.BL_RodzajTransakcji where MetodaPlatnosciBL = '" + cumulative_response.Orders[order_nr].PaymentMethod + "')) " +
                "begin if @xlkod IS NOT NULL " +
                "update tp set tp.TrP_KARNumer = @xlkod from cdn.TraPlat as i join cdn.TraPlat as tp on i.TrP_GIDNumer = tp.TrP_GIDNumer and i.TrP_GIDTyp = tp.TrP_GIDTyp " +
                "join cdn.TraNag as t on t.TrN_GIDNumer = tp.TrP_GIDNumer and t.TrN_GIDTyp = tp.TrP_GIDTyp and tp.TrP_FormaNazwa = @rodzpl " +
                "where t.TrN_ZaNNumer = (select ZaN_GIDNumer from cdn.ZamNag where ZaN_DokumentObcy = @esklep) " +
                "end " +
                "begin update t set t.TrN_KARNumer = tp.TrP_KARNumer from cdn.TraNag as i join cdn.TraNag as t on i.TrN_GIDNumer = t.TrN_GIDNumer and i.TrN_GIDTyp = t.TrN_GIDTyp " +
                "join cdn.TraPlat as tp on t.TrN_GIDNumer = tp.TrP_GIDNumer and t.TrN_GIDTyp = tp.TrP_GIDTyp and tp.TrP_FormaNazwa = @rodzpl " +
                "where t.TrN_ZaNNumer = (select ZaN_GIDNumer from cdn.ZamNag where ZaN_DokumentObcy = @esklep) " +
                "end ";
            Console.WriteLine(ConnectToSQL(connection_string, karNr));
            Console.WriteLine(
                ConnectToSQL(connection_string,
            "declare @kntgid as varchar (128), @zangid as varchar (128), @zantyp as varchar(128), @trngid as varchar(128), @trntyp as varchar(128) " +
            "set @kntgid = (select top 1 KnT_GIDnumer from cdn.KntKarty where KnT_Akronim=('KNT_'+'"+kontrahent[2]+ "')) " +
            "set @zangid = (select top 1 ZaN_GIDNumer from cdn.ZamNag where ZaN_KntNumer = @kntgid) " +
            "set @zantyp = (select top 1 ZaN_GIDTyp from cdn.ZamNag where ZaN_KntNumer = @kntgid) " +
            "set @trntyp = 1824 " +
            "set @trngid = (select top 1 TrN_GIDNumer from cdn.TraNag where TrN_KntNumer = @kntgid and TrN_GIDTyp=@trntyp) " +
            "update cdn.TraNag set TrN_FrmNumer = (" +
            "select ZaN_FrmNumer from cdn.ZamNag as t join cdn.FrmStruktura as s on t.ZaN_FrSID = s.FRS_ID " +
            "join cdn.Firma as f on s.FRS_FrMNumer = f.FRM_GIDNumer " +
            "where ZaN_GIDTyp = @zantyp AND ZaN_GIDFirma = 1034247 AND ZaN_GIDNumer = @zangid) " +
            "where TrN_GIDTyp = @trntyp AND TrN_GIDFirma = 1034247 AND TrN_GIDNumer = @trngid"));
            Console.WriteLine(
            ConnectToSQL(connection_string,
            $@"declare  @netto as decimal 
                            declare  @gid as int 
                            declare  @typ as int 
                            set @gid = {xLDokumentNagInfo.GIDNumer} 
                            set @typ = {xLDokumentNagInfo.GIDTyp} 
                            set @netto = (select SUM(TRV_NettoR) from cdn.traVat where TrV_GIDNumer=@gid and TrV_GIDTyp=@typ) 
                            if (@netto<=0) 
                            begin
                            update cdn.TraNag set TrN_Rozliczac=0 where  TrN_GIDNumer=@gid and TrN_GIDTyp=@typ 
                            update cdn.TraVat set TrV_Rozliczac=0 where  TrV_GIDNumer=@gid and TrV_GIDTyp=@typ end"));
            return return_value;
        }

        public void AddFS_FROM_ZS(TextWriter tw_log)
        {
            
            string fsList = ConnectToSQL(connection_string, $@"Declare @Numbers AS varchar(max) 
            SELECT  @Numbers = COALESCE(@Numbers + ';', '') + CAST(ZaN_GIDNumer AS varchar(256)) 
			from(
            select sum(ZaE_Ilosc) as ZamIlosc  
            , sum(isnull(TwZ_Ilosc, 0)) as Zasob 
            , ZaN_GIDNumer 
            , ZaE_GIDLp 
            from cdn.ZamNag as z 
            join cdn.ZamElem as e on z.ZaN_GIDNumer = e.ZaE_GIDNumer and z.ZaN_GIDTyp = e.ZaE_GIDTyp 
            join cdn.TwrKarty as t on e.ZaE_TwrNumer = t.Twr_GIDNumer and e.ZaE_TwrTyp = t.Twr_GIDTyp and t.Twr_Typ = 1 
            join cdn.KntKarty on ZaN_KntNumer = Knt_GIDNumer and ZaN_KntTyp = Knt_GIDTyp and Knt_Akronim like 'KNT_%' 
            outer apply (select sum(TwZ_Ilosc) as TwZ_Ilosc from cdn.TwrZasoby as tz where e.ZaE_TwrNumer = tz.TwZ_TwrNumer and e.ZaE_TwrTyp = tz.TwZ_TwrTyp) as za 
            where ZaN_Stan between 3 and 6 
            --and ZaN_GIDNumer = 52670 
            group by ZaN_GIDNumer, ZaE_GIDLp
            having sum(ZaE_Ilosc) <= sum(za.TwZ_Ilosc) 
            ) as x 
            group by ZaN_GIDNumer 
            having sum(ZamIlosc) <= sum(Zasob) 
                      SELECT  @Numbers");
            string[] separator = new string[1] { ";" };
            string[] orderList = fsList.Split(separator, StringSplitOptions.RemoveEmptyEntries);
            Console.WriteLine(fsList);
            foreach (var order in orderList)
            {
                CreateFS(tw_log, order);
            }
        }
        public void CreateFS(TextWriter tw_log, string order)
        {

            AttachThreadToClarion(true);
            int XLFS_id = 1;
            int return_value = 0;
            XLZwiazanyDokInfo_20220 CzyszczenieListy = new XLZwiazanyDokInfo_20220
            {
                Wersja = XLWersjaAPI,
                GIDTyp = 32,
                GIDNumer = 0,
                Akcja = -1,
                Ilosc = "1",

            };
            return_value = cdn_api.cdn_api.XLDokumentyZwiazane(CzyszczenieListy);
            Console.WriteLine($@"Czysczenie listy, kod api: {return_value}");
            string czesciowaRealizacja = ConnectToSQL(connection_string, $@"select top 1 Atr_Wartosc from cdn.AtrybutyKlasy left join cdn.Atrybuty on AtK_ID = Atr_AtkId left join cdn.AtrybutyHist on AtH_AtkId = Atr_AtkId where AtK_Nazwa = 'Czesciowa Realizacja' and Atr_ObiNumer ={order}");
            int tryb = 0;
            string doesExist = ConnectToSQL(connection_string, $@"if exists(select top 1 TrN_KntNumer from cdn.TraNag where TrN_GIDTyp=2033 and TrN_ZaNNumer={order}) 
                select top 1 1 
                else
                if not exists (select Knt_Akronim from cdn.KntKarty where Knt_GIDNumer=(select ZaN_KntNumer from cdn.ZamNag where ZaN_GIDNumer={order}) and Knt_Akronim LIKE 'KNT_%')
                select top 1 1
                else 
                select case when 
                (select Sum(ZaE_Ilosc) from cdn.ZamElem where ZaE_GIDNumer={order} and ZaE_TwrKod!='US.TRANSP.' )<=
                (select sum(ZaE_Ilosc) from cdn.TwrZasoby left join cdn.ZamElem on TwZ_TwrNumer=ZaE_TwrNumer where ZaE_GIDNumer={order})
                then 0
                else 2 end");
            Console.WriteLine(doesExist);
            var warunek = "";
            //warunek = ConnectToSQL(connection_string,
            //$@"select CzyFSL from ord.BL_Zrodla where WartoscXL= 
            //         (select top 1 Atr_Wartosc from cdn.AtrybutyKlasy left join cdn.Atrybuty on AtK_ID = Atr_AtkId left join cdn.AtrybutyHist on AtH_AtkId = Atr_AtkId where AtK_Nazwa = 'Źródło' and Atr_ObiNumer = {order})")) != "1")
           
            if (doesExist == "1" ||
                doesExist == "2" ||
                czesciowaRealizacja == "TAK") 
                {
                if(doesExist!="2")tw_log.WriteLine(DateTime.Now + $@" Dodawanie FS wstrzymano na zamówieniu {order} z kodem {doesExist} {"\n"} Czesc. Realizacja {czesciowaRealizacja} z warunkiem {warunek} " );
                return;
                }

            Console.WriteLine("Dodanie FS do zamowienia nr " + order);
            tw_log.WriteLine(DateTime.Now + " Dodanie FS do zamowienia nr " + order);
            string cecha = ConnectToSQL(connection_string, "select top 1 ZaN_CechaOpis from cdn.ZamNag where ZaN_GIDNumer = " + order);
            string mag = ConnectToSQL(connection_string, "Select BL_Zrodla.SerieZS" +
            " from ORD.BL_Zrodla where " +
            "BL_Zrodla.WartoscBL =" + cecha);
            string kraj = ConnectToSQL(connection_string, $@"select ZaN_KrajPrzezWys from cdn.ZamNag where ZaN_GIDNumer={order}");
            string opis = ConnectToSQL(connection_string, $@"select top 1 ZnO_Opis from cdn.ZaNOpisy where ZnO_ZamNumer={order}");
            string spDostawy = ConnectToSQL(connection_string, $@"select ZaN_SpDostawy from cdn.ZamNag where ZaN_GIDNumer={order}");
            string czyVat = "0";
            string qry = $@"select MAX(TwZ_DataP)
                            from cdn.TwrZasoby
                            /*left join cdn.TwrKarty on TwZ_TwrNumer = Twr_GIDNumer */
                            left join cdn.ZamElem on TwZ_TwrNumer = ZaE_TwrNumer
                            where ZaE_GIDNumer = {order}
                            and TwZ_Ilosc > 0 ";

            string dataSprMag = ConnectToSQL(connection_string, qry);
            if (Double.TryParse(dataSprMag, out double result)!=true)
            { 
                tw_log.WriteLine(DateTime.Now + $@" Zamówienie o GID {order} zwróciło NULL po zapytaniu o datę przyjęcia towaru");
                return;
            }
            string zw_Dok = ConnectToSQL(connection_string, $@"
                    declare @results varchar(512) 
                    select @results = coalesce(@results + ';', '') + convert(varchar(128), TrN_GIDNumer) 
                    from cdn.TraNag 
                    where TrN_ZaNNumer = {order} and TrN_GIDTyp = 1824 
                    select @results as results");
            int count = 0;
            if (zw_Dok!="")
            {
                try
                {
                    string[] zw_Dok_Arr = zw_Dok.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToArray();
                   

                    foreach (var dokument in zw_Dok_Arr)
                    {
                        count++;
                            if (Int32.TryParse(dokument, out int dokument_gid))
                            {
                               
                                XLZwiazanyDokInfo_20220 xLZwiazanyDokInfo = new XLZwiazanyDokInfo_20220
                                {
                                    Wersja = XLWersjaAPI,
                                    GIDTyp = 1824,
                                    GIDNumer = dokument_gid,
                                    Akcja = 0,
                                    Ilosc = "1",
                                };
                                return_value = cdn_api.cdn_api.XLDokumentyZwiazane(xLZwiazanyDokInfo);
                                tw_log.WriteLine(DateTime.Now + $" Dodano {count}. dokument zwiazany  o gid " + dokument + " zwrócony kod API: " + return_value);
                                Console.WriteLine($"Pre 9A.2 Dodano {count}. dokument zwiazany o gid " + dokument + " zwrócony kod API: " + return_value); 
                            } else
                            {
                                Console.WriteLine(DateTime.Now + $@" Nie dodano dok. zwiazanego dla , SQL zwraca: {zw_Dok} wartosc: {dokument} ");
                                tw_log.WriteLine(DateTime.Now + $@" Nie dodano dok. zwiazanego dla , SQL zwraca: {zw_Dok} wartosc: {dokument} ");
                                return;
                            }  
                    }
                }
                catch (Exception)
                {
                    if (count>0)
                    {
                        XLZwiazanyDokInfo_20220 xLZwiazanyDokInfo = new XLZwiazanyDokInfo_20220
                        {
                            Wersja = XLWersjaAPI,
                            GIDTyp = 1824,
                            GIDNumer = 0,
                            Akcja = -1,
                            Ilosc = "1",
                        };
                        return_value = cdn_api.cdn_api.XLDokumentyZwiazane(xLZwiazanyDokInfo);
                        tw_log.WriteLine(DateTime.Now + " Nie dodano dok. zwiazanego, SQL zwraca: " + zw_Dok + " zwrócony kod API: " + return_value);
                    }

                    return;
                }
            }

            

            Console.WriteLine(qry + "\n" + dataSprMag + "to data przyjecia na mag towaru " + order);
            //Console.ReadKey();
            /*
            Console.WriteLine("od 1990 " + dataSprMag);
            Console.WriteLine("unix timestamp " + (double.Parse(dataSprMag) + 630720000 + 432000).ToString());
            Console.WriteLine("Clarion timestamp " + (GetClarionDate
                        (UnixTimeStampToDateTime
                            (double.Parse(dataSprMag) + 630720000 + 432000)).ToString()));
            */
            string sNumbers = ConnectToSQL(connection_string, "SELECT top 1 Concat(ISNULL(CAST(ZaN_KntTyp as varchar), ''), ';', ISNULL(CAST(ZaN_KntFirma as varchar), ''), ';', ISNULL(CAST(ZaN_KntNumer as varchar), ''), ';', ISNULL(CAST(ZaN_KntLp as varchar), ''), ';'," +
            "ISNULL(CAST(ZaN_GIDTyp as varchar), ''), ';', ISNULL(CAST(ZaN_GIDFirma as varchar), ''), ';', ISNULL(CAST(ZaN_GIDNumer as varchar), ''), ';' + ISNULL(CAST(ZaN_GIDLp as varchar), ''),';', ISNULL(CAST(ZaN_FormaNr as varchar), '')) as Odp" +
            " FROM cdn.ZamNag where ZaN_GIDNumer = " + order);

            if (sNumbers.Length > 0)
            {
                var numbers = sNumbers.Split(';').Select(Int32.Parse).ToList();
                Console.WriteLine($"GID {numbers[0]}, {numbers[1]}, {numbers[2]}, {numbers[3]}");
                XLDokumentNagInfo_20220 xLDokumentNagInfo = new XLDokumentNagInfo_20220
                {
                    Wersja = XLWersjaAPI,
                    Typ = 2033,
                    GIDNumer = 0,
                    GIDFirma = 0,
                    GIDLp = 0,
                    GIDTyp = 0,
                    KntTyp = numbers[0],
                    KntFirma = numbers[1],
                    KntNumer = numbers[2],
                    KntLp = numbers[3],
                    ZamTyp = numbers[4],
                    ZamFirma = numbers[5],
                    ZamNumer = numbers[6],
                    ZamLp = numbers[7],
                    Forma = numbers[8],
                    Tryb = 2,
                    Rejestr = "SDET",
                    NB = "B",
                    Rok = DateTime.Now.Year,
                    Miesiac = DateTime.Now.Month,
                    MiesiacVAT = DateTime.Now.Month,
                    RokVAT = DateTime.Now.Year,
                    DataSpr =
                    Int32.Parse
                        (GetClarionDate
                            (UnixTimeStampToDateTime
                                (double.Parse(dataSprMag) + 630720000 + 432000)).ToString()),
                    DataMag =
                    Int32.Parse
                        (GetClarionDate
                            (UnixTimeStampToDateTime
                                (double.Parse(dataSprMag) + 630720000 + 432000)).ToString()),
                    SposobDst = spDostawy,
                    Kraj = kraj,
                    Zwiazane = 1,
                    Cecha = cecha,
                    Opis = opis,
                    //IgnorujMagazynowe = 2,
                    //IgnorujZaliczkowe = 2,
                    //FRSID = 1,
                };
                if (ConnectToSQL(connection_string, $@"select cast((TrN_NettoR+TrN_VatR)-(TrV_NettoR+TrV_VatR) as float) from cdn.TraNag left join cdn.TraVat on TrN_GIDNumer=TrV_GIDNumer and TrN_GIDTyp=TrV_GIDTyp where TrN_GIDNumer=(select TrN_GIDNumer from cdn.TraNag where TrN_KntNumer={numbers[2]}) and TrN_GIDTyp=1824)") == "0")
                {
                    xLDokumentNagInfo.DeklaracjaVAT7 = 2;
                    xLDokumentNagInfo.Rozliczac = 0;
                    czyVat = "2";
                }
                if(ConnectToSQL(connection_string, $@"select top 1 Knt_PlatnikVat from Cdn.KntKarty where Knt_GIDNumer={xLDokumentNagInfo.KntNumer}") =="0")
                {
                    xLDokumentNagInfo.Fiskalny = 1;
                }
                try
                {
                    return_value = cdn_api.cdn_api.XLNowyDokument(sesja, ref XLFS_id, xLDokumentNagInfo);
                }
                catch (Exception e)
                {
                    throw new System.Exception("Wystapil wyjatek przy dodawaniu faktury. Kod API: " + return_value + ", Wyjątek: " + e);
                }
                finally
                {
                    Console.WriteLine("9A. API po dodaniu faktury zwrocilo : " + return_value);
                }
                tw_log.WriteLine($@"3. Dodanie FS o gid {xLDokumentNagInfo.GIDNumer}. API zwróciło kod: " + return_value);
                if (return_value != 0)
                {
                   
                    tryb = -1;
                }

                //Ilosc Elem. zam. 
                string elem = ConnectToSQL(connection_string, $@"select COUNT(*) from cdn.ZamElem where ZaE_GIDNumer={order}");
                Console.WriteLine("elem : " + elem);
                if (Int32.TryParse(elem, out int res) == true)
                {
                    for (int product_nr = 1; product_nr <= res; product_nr++)
                    {
                        string Zam_Info = ConnectToSQL(connection_string, $@"select CONCAT(ZaE_Ilosc, '#', ZaE_TwrKod,'#', ZaE_TwrNazwa,'#', ZaE_CenaUzgodniona,'#', ZaE_FlagaVat,'#',
                            ZaE_StawkaPod,'#', ZaE_GIDTyp,'#',ZaE_GIDNumer,'#',ZaE_GIDFirma,'#',ZaE_GrupaPod,'#',ZaE_Waluta) 
                            from cdn.ZamElem as c where ZaE_GIDNumer={order} and ZaE_GIDLp={product_nr}");
                        string[] Zam_Info_Arr = Zam_Info.Split('#');
                        if (Zam_Info_Arr.Length >= 10)
                        {
                            Console.WriteLine("Dodawanie towaru o kodzie " + Zam_Info_Arr[1]);
                            string ilosc = "SELECT ISNULL(SUM([TwZ_Ilosc]),0) FROM CDN.TwrZasoby join cdn.TwrKarty on TwZ_TwrTyp=Twr_GIDTyp and TwZ_TwrFirma=Twr_GIDFirma and  TwZ_TwrNumer=Twr_GIDNumer and TwZ_TwrLp=Twr_GIDLp"
                                            + " WHERE Twr_Kod='" + Zam_Info_Arr[2] + "'";
                            tw_log.WriteLine("Dostępna ilość: "+(ConnectToSQL(connection_string, ilosc)));
                            XLDokumentElemInfo_20220 xLDokumentElemInfo = new XLDokumentElemInfo_20220
                            {
                                Wersja = XLWersjaAPI,
                                Ilosc = Zam_Info_Arr[0],
                                TowarKod = Zam_Info_Arr[1],
                                TowarNazwa = Zam_Info_Arr[2],
                                Cena = Zam_Info_Arr[3],
                                PrecyzjaCeny = 2,
                                FlagaVAT = Int32.Parse(Zam_Info_Arr[4]),
                                StawkaPod = Int32.Parse((Zam_Info_Arr[5]).Replace(".", "")),
                                ZamTyp = Int32.Parse(Zam_Info_Arr[6]),
                                ZamNumer = Int32.Parse(Zam_Info_Arr[7]),
                                ZamFirma = Int32.Parse(Zam_Info_Arr[8]),
                                ZamLp = product_nr,
                                Vat = Zam_Info_Arr[9],
                                Kraj = kraj,
                                Waluta = Zam_Info_Arr[10],
                                TwrReq = 0,
                                JmZ = (Zam_Info_Arr[1] == "US.TRANSP.") ? "opł" : "szt",
                                //PrzeliczL = 1,F
                                //NiePrzeliczaj=1,
                            };
                            if (Zam_Info_Arr[1] != "US.TRANSP.") xLDokumentElemInfo.Magazyn = "M" + mag;
                            try
                            {
                                return_value = cdn_api.cdn_api.XLDodajPozycje(XLFS_id, xLDokumentElemInfo);
                            }
                            catch (Exception e)
                            {
                                throw new System.Exception("Wystapil wyjatek przy dodawaniu elementu. Kod API: " + return_value + ", Wyjątek: " + e);
                            }
                            finally
                            {
                                Console.WriteLine("9B. API po dodaniu elementu zwrocilo : " + return_value + " flaga: " + xLDokumentElemInfo.Flagi);
                            }
                            if (return_value != 0)
                            {
                                tw_log.WriteLine("3B Błąd przy dodaniu elementu do faktury, nr API to :" + return_value);
                                tryb = -1;
                            }
                        }
                        else
                        {
                            tw_log.WriteLine("Element nie poziada wystarczajacej liczby pozycji!");
                            tw_log.Write($"Info do elementu zamówienia: {Zam_Info}");
                            tryb = -1;
                        }

                    }

                }

                XLZamkniecieDokumentuInfo_20220 xLZamkniecieDokumentuInfo = new XLZamkniecieDokumentuInfo_20220
                {
                    Wersja = XLWersjaAPI,
                    GidNumer = xLDokumentNagInfo.GIDNumer,
                    Tryb = tryb,
                    Magazynowe = 3,
                    LimitKnt = 2,
                };
                if (czyVat == "2")
                {
                    xLZamkniecieDokumentuInfo.Fiskalny = 2;
                }
                //nie wystawiają się WM z jakiegoś powodu
                //else xLZamkniecieDokumentuInfo.Fiskalny = 0;

                try
                {
                    return_value = cdn_api.cdn_api.XLZamknijDokument(XLFS_id, xLZamkniecieDokumentuInfo);
                }
                catch (Exception e)
                {
                    throw new System.Exception("Wystapil wyjatek przy dodawaniu faktury. Kod API: " + return_value + ", Wyjątek: " + e);
                }
                finally
                {

                    Console.WriteLine("9C. API po zamknieciu faktury zwrocilo : " + return_value);
                    tw_log.WriteLine("3C. API po zamknieciu faktury zwrocilo : " + return_value);
                }

                /*
                try
                {
                    XLDokumentMagNagInfo_20220 wm = new XLDokumentMagNagInfo_20220();
                    wm.Typ = 1601;
                    wm.Wersja = XLWersjaAPI;
                    wm.ZrdFirma = xLDokumentNagInfo.GIDFirma;
                    wm.ZrdLp = 1; //xLDokumentNagInfo.GIDLp;
                    wm.ZrdNumer = xLDokumentNagInfo.GIDNumer;
                    wm.ZrdTyp = xLDokumentNagInfo.GIDTyp;
                    int wmID = 0;
                    int bladWM = cdn_api.cdn_api.XLNowyDokumentMag(sesja, ref wmID, wm);

                    if (bladWM != 0)
                        tw_log.WriteLine($"Błąd dodawania WM: {bladWM}");

                    XLZamkniecieDokumentuMagInfo_20220 wmZam = new XLZamkniecieDokumentuMagInfo_20220();
                    wmZam.Wersja = XLWersjaAPI;
                    wmZam.Tryb = 0;
                    bladWM = cdn_api.cdn_api.XLZamknijDokumentMag(wmID, wmZam);

                    if (bladWM != 0)
                        tw_log.WriteLine($"Błąd zamykania WM: {bladWM}");

                }
                catch (Exception e)
                {
                    tw_log.WriteLine($"Dodawanie WM: {e.Message}");
                }*/

                Console.WriteLine("GID FAKTURY TO : " + xLDokumentNagInfo.GIDNumer);
                Console.WriteLine(
                    ConnectToSQL(connection_string,
                        "declare @kntgid as varchar (128), @zangid as varchar (128), @zantyp as varchar(128), @trngid as varchar(128), @trntyp as varchar(128) " +
                        "set @kntgid = (select top 1 KnT_GIDnumer from cdn.KntKarty where KnT_Akronim=('KNT_'+'" + numbers[2] + "')) " +
                        "set @zangid = (select top 1 ZaN_GIDNumer from cdn.ZamNag where ZaN_KntNumer = @kntgid) " +
                        "set @zantyp = (select top 1 ZaN_GIDTyp from cdn.ZamNag where ZaN_KntNumer = @kntgid) " +
                        "set @trntyp = 2033 " +
                        "set @trngid = (select MAX(TrN_GIDNumer) from cdn.TraNag where TrN_KntNumer = @kntgid and TrN_GIDTyp=@trntyp) " +
                        "update cdn.TraNag set TrN_FrmNumer = (" +
                        "select top 1 ZaN_FrmNumer from cdn.ZamNag as t join cdn.FrmStruktura as s on t.ZaN_FrSID = s.FRS_ID " +
                        "join cdn.Firma as f on s.FRS_FrMNumer = f.FRM_GIDNumer " +
                        "where ZaN_GIDTyp = @zantyp AND ZaN_GIDFirma = 1034247 AND ZaN_GIDNumer = @zangid) " +
                        "where TrN_GIDTyp = @trntyp AND TrN_GIDFirma = 1034247 AND TrN_GIDNumer = @trngid"));
                Console.WriteLine(
                    ConnectToSQL(connection_string,
                    "update cdn.TraNag set TrN_FrsID = (select IdSklepuERP from ord.BL_Zrodla where WartoscXL ='" + cecha + "') where TrN_KntNumer = " + numbers[2] + " and TrN_GIDTyp = 2033"));
                Console.WriteLine(
                    ConnectToSQL(connection_string,
                    $@"declare  @netto as decimal 
                    declare  @gid as int 
                    declare  @typ as int 
                    set @gid = {xLDokumentNagInfo.GIDNumer} 
                    set @typ = {xLDokumentNagInfo.GIDTyp} 
                    set @netto = (select SUM(TRV_NettoR) from cdn.traVat where TrV_GIDNumer=@gid and TrV_GIDTyp=@typ) 
                    if (@netto<=0) begin 
                    update cdn.TraNag set TrN_Rozliczac=0 where  TrN_GIDNumer=@gid and TrN_GIDTyp=@typ 
                    update cdn.TraVat set TrV_Rozliczac=0 where  TrV_GIDNumer=@gid and TrV_GIDTyp=@typ end "));
            }
        }

        string WyliczSume (string ean)
        {
            List<int> codeIntList = new List<int>();

            //String finalEan;
            int sumOddNumbers = 0;
            int sumEvenNumbers = 0;
            int resultOddPlusEven = 0;
           // int saveLastNumber = 0;

            //String eanText = howLongIsEan(ean);
            int eanIsLong = ean.Length;

           
                for (int i = 0; i < eanIsLong; i++)
                {
                    codeIntList.Add(Convert.ToInt32(ean[i]));
                }

            //saveLastNumber = codeIntList.Get(codeIntList.Size() - 1);
            //codeIntList.remove(codeIntList.size() - 1);
            //System.Collections.reverse(codeIntList);

                for (int i = 0; i < codeIntList.Count() ; i++)
                {
                    if (i % 2 == 0)
                    {
                        sumOddNumbers += codeIntList[i];
                    }
                    else
                    {
                        sumEvenNumbers += codeIntList[i]*3;
                    }
                }
                resultOddPlusEven = sumOddNumbers + sumEvenNumbers;
            

            int modulo = 10 - (resultOddPlusEven % 10);
            return Convert.ToString(modulo);
        }
            static async Task Main()
        {
            //Base64Encoder pdfencoder = new Base64Encoder(); 
            //AttachThreadToClarion(1);
                        Console.WriteLine("Wczytywanie konfiguracji");
            if (program.Get_Config())
            {
                if (program.test == 1)
                {
                    program.XLDatabase = "DOMAR2";
                    program.XLUser = "ADMIN";
                    program.XLPass = "NE!w0rk12";
                    program.XLKey = "SRV1::5000139268";
                    program.XLConString = ";Integrated Security=SSPI;";
                }
                Console.WriteLine("Logowanie do ERP XL...");
                int num = program.Login_XL();
                if (num == 0)
                {
               
                    Console.WriteLine("Connection string ...");
                    Console.WriteLine(program.GetConnectionString());
                    try
                    {
                       await Baselinker_Connect(program.days);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Wystąpił bład " + e.ToString());
                    }
                    finally
                    {
                        try { 
                       await CreateIndex_XL();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Wystąpił bład CreateIndex " + e.ToString());
                            //Console.ReadKey();
                        }
                        //AddFS_FROM_ZS();
                    }
                    
                }
                else
                { 
                    Console.WriteLine("Nie udało się zalogować do XL. Praca zostanie zakończona. Błąd logowania nr: " + num.ToString());
                   //Console.ReadKey();
                }
            }
            else Console.WriteLine("Nie udało się odczytać danych z pliku konfiguracjnego.");
            Console.WriteLine("ERPXL Wylogowano!");
           // Console.ReadKey();
            Logout_XL();
        }


        //string json = System.IO.File.ReadAllText(@"C:\Users\krzysztof.bouzyk\source\repos\OneS-Baselinker\OneS-Baselinker\WriteLines.txt");
    }
}
