using cdn_api;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportFW
{
    public partial class Form1 : Form
    {
        private static int sesja = 0;
        private static int XLWersjaAPI = 20220;
        private static int XLTrybWsadowy = 1;
        private static int XLWinieta = -1;
        private static int XLUtworzWlasnaSesje = 1;
        private static string XLSQLServer = "172.51.1.8";
        private static string XLSQLDatabase = "ElektroKlim";
        private static string XLUser = "ADMIN"; //
        private static string XLPass = "NE!w0rk12";//"Dupa2003!";
        private static int dokID = 0;
        string connectionString = @"Data Source = EKSRV04\COMARCH; Initial Catalog = ERPXL_E; User ID = ComarchAPI; Password=hT@S+NQZBT";
        string date1 = "1800-12-28";

        //@"Data Source = SQLSERVER; Initial Catalog = DomarXL; User ID=AdminS;Password=6,eF#W451,t6,D@";

        public Form1()
        {
            //this.connectionString = GetConnectionString();
            InitializeComponent();
        }


        void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                #region excel5
                #region wybranie pliku
                Cursor.Current = Cursors.WaitCursor;
                if (tB1.Text == "nie wybrano pliku")
                {
                    MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string filePath = tB1.Text;
                DataTable dataTable;
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader;
                    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = false//true
                        }
                    };

                    var dataSet = reader.AsDataSet(conf);
                    dataTable = dataSet.Tables[0];
                }
                #endregion

                #region Przypisywanie danych z datasetu do listy przyszłych pozycji.

                SprawdzamCzyArkuszNieJestPusty(dataTable);
                dodajDokumenty(dataTable);

                #endregion

                #endregion
            }
            catch(Exception en)
            {
                MessageBox.Show("Błąd pobierania danych z excel: " + en.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }
        #region Dodawanie dokumentów
        void dodajDokumenty(DataTable dataTable)
        {
            bool jest = true;
            for(int k=1; k<dataTable.Rows.Count; k++)
            {
                try
                {
                    //Sprawdzenie, czy kontrahent jest w bazie
                    //PYTANIE!!! Zawsze będzie prfiks ' ' nip ? Co gdy kontrahent jest polski?
                    if (Convert.ToString(dataTable.Rows[k][10]).Trim() != "")
                    {
                        string nip = Convert.ToString(dataTable.Rows[k][10]).Trim();
                        string jestQuery = "if exists (SELECT * FROM [CDN].[KntKarty] WHERE Knt_NipPrefiks+' '+Knt_NipE='" + nip + "') SELECT 'TAK' else SELECT 'NIE'";
                        string istnieje = ConnectToSQL(connectionString, jestQuery);
                        if (istnieje == "NIE")
                        {
                            MessageBox.Show("Brak kontrahenta o NIPie: " + nip + "\nPrzerwano działanie programu", "Komunikat",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            jest = false;
                            return;
                        }
                    }
                }
                catch(Exception e)
                {
                    MessageBox.Show("Błąd sprawdzania kontrahenta: " + e.Message+ "\nPrzerwano działanie programu", "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning) ;
                    jest = false;
                    return;
                }

            }

            if(jest==false)
            {
                return;
            }

            XLLoginInfo_20220 nowyUzytkownik = new XLLoginInfo_20220();
            nowyUzytkownik.Wersja = XLWersjaAPI;
            nowyUzytkownik.Winieta = XLWinieta;
            nowyUzytkownik.TrybWsadowy = XLTrybWsadowy;
            nowyUzytkownik.UtworzWlasnaSesje = XLUtworzWlasnaSesje;
            nowyUzytkownik.SerwerKlucza = XLSQLServer;
            nowyUzytkownik.Baza = XLSQLDatabase;
            nowyUzytkownik.OpeIdent = XLUser;
            nowyUzytkownik.OpeHaslo = XLPass;
            nowyUzytkownik.PortHTTPSerweraKlucza = 1;
            nowyUzytkownik.RodzajSerweraKlucza = 0;
            nowyUzytkownik.SerwerKlucza = @"EKSRV04\COMARCH::5000235184";
            nowyUzytkownik.ProgramID = "ONE-027 Import FW";

            //MessageBox.Show(nowyUzytkownik.SerwerKlucza);

            int blad = cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);

            if (blad == 0)
            {
                //MessageBox.Show(nowyUzytkownik.SerwerKlucza);
                pBar1.Visible = true;
                // Set Minimum to 1 to represent the first file being copied.
                pBar1.Minimum = 1;
                // Set Maximum to the total number of files to copy.
                pBar1.Maximum = dataTable.Rows.Count;
                // Set the initial value of the ProgressBar.
                pBar1.Value = 1;
                // Set the Step property to a value of 1 to represent each file being copied.
                pBar1.Step = 1;

                for (int k = 1; k < dataTable.Rows.Count; k++)
                {
                    pBar1.PerformStep();
                    try
                    { 
                        if(Convert.ToString(dataTable.Rows[k][10]).Trim()!="")
                            dodajFW(dataTable, k, sesja);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Błąd wywoływania funkcji dla faktury " + Convert.ToString(dataTable.Rows[k][4]).Trim() + ": " + e.Message, "Komunikat",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                pBar1.Value = 1;
                pBar1.Visible = false;
                
                cdn_api.cdn_api.XLLogout(sesja);


            }
            else
            {
                MessageBox.Show("Błąd logowania:  " + blad, "Komunikat",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }





            MessageBox.Show("Zakończono dodawanie dokumentów FW", "Komunikat",
                                     MessageBoxButtons.OK, MessageBoxIcon.Information);
        }



        public void dodajFW(DataTable dataTable, int k, int sesja)
        {
            string date2 = "";
            if (Convert.ToString(dataTable.Rows[k][2]).IndexOf(".") == 1)
                date2 = Convert.ToString(dataTable.Rows[k][2]).Substring(5, 4) + "-" + Convert.ToString(dataTable.Rows[k][2]).Substring(2, 2) + "-0" + Convert.ToString(dataTable.Rows[k][2]).Substring(0, 1);
            else
                date2 = Convert.ToString(dataTable.Rows[k][2]).Substring(6, 4) + "-" + Convert.ToString(dataTable.Rows[k][2]).Substring(3, 2) + "-" + Convert.ToString(dataTable.Rows[k][2]).Substring(0, 2);

            string adrNumerQuery = "SELECT TOP(1) [KnA_GIDNumer] FROM [CDN].[KntAdresy] WHERE KnA_Akronim = "
                                    + " (SELECT TOP(1) Knt_Akronim FROM [CDN].[KntKarty] WHERE Knt_NipPrefiks + ' ' + Knt_NipE = '" + Convert.ToString(dataTable.Rows[k][10]).Trim() + "') and KnA_DataArc = 0";
            int adrNumer = Convert.ToInt32(ConnectToSQL(connectionString, adrNumerQuery));

            int dokID = 0;
            XLDokumentNagInfo_20220 dok = new XLDokumentNagInfo_20220();
            dok.Wersja = XLWersjaAPI;
            dok.GIDFirma = 0;
            dok.GIDLp = 0;
            dok.GIDNumer = 0;
            dok.Avista = 1;
            dok.Data= Convert.ToInt32(ConnectToSQL(connectionString, "SELECT DATEDIFF(day, '" + date1 + "','" + date2 + "'); "));
            dok.DataSpr= Convert.ToInt32(ConnectToSQL(connectionString, "SELECT DATEDIFF(day, '" + date1 + "','" + date2 + "'); "));
            dok.DataVat= Convert.ToInt32(ConnectToSQL(connectionString, "SELECT DATEDIFF(day, '" + date1 + "','" + date2 + "'); "));
            dok.DokumentObcy = Convert.ToString(dataTable.Rows[k][4]);
            dok.Typ = 2036;
            dok.Akronim = ConnectToSQL(connectionString, "SELECT TOP(1) Knt_Akronim FROM [CDN].[KntKarty] WHERE Knt_NipPrefiks + ' ' + Knt_NipE = '" + Convert.ToString(dataTable.Rows[k][10]).Trim() + "'");
            dok.IncotermsSymbol = "EXW";
            dok.KodRodzajuTransakcji = "11";
            dok.Opis= Convert.ToString(dataTable.Rows[k][7]);
            dok.Tryb = 2;
            dok.AdrNumer = adrNumer;
            dok.AdrTyp = 864;
            dok.AdrLp=0;
            dok.AdrFirma = 1475073;
            dok.Rejestr = "SPRWW";
            dok.ExpoNorm = 6;

            int blad = cdn_api.cdn_api.XLNowyDokument(sesja, ref dokID, dok);
            if(blad!=0)
            {
                MessageBox.Show("Błąd dodawania dokumentu "+ Convert.ToString(dataTable.Rows[k][4]) +":  " + blad, "Komunikat",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;

            }


            XLVatInfo_20220 vat = new XLVatInfo_20220();
            vat.DeklMiesiac = Convert.ToDateTime(date2).Month;
            vat.DeklRok = Convert.ToDateTime(date2).Year;
            vat.IgnorujVat = 1;
            vat.RodzajZakupu = 1;
            vat.Rozliczac = 1;
            vat.RozliczacP = 1;
            vat.StawkaPod = 0;
            vat.WartoscVat = "0" ;
            vat.FlagaVat = 1;
            vat.Wartosc= Math.Round(Convert.ToDouble(dataTable.Rows[k][6]), 2, MidpointRounding.AwayFromZero).ToString();
            vat.Wersja = XLWersjaAPI;
            vat.Struktura = 1;
            blad = cdn_api.cdn_api.XLDodajVAT(dokID, vat);
            if (blad != 0)
            {
                MessageBox.Show("Błąd dodawania vatu do dokumentu " + Convert.ToString(dataTable.Rows[k][4]) + ":  " + blad, "Komunikat",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;

            }

           /* XLJPKV7Info_20220 jpk = new XLJPKV7Info_20220();
            jpk.Wersja = XLWersjaAPI;
            //jpk.TrNFirma = dok.GIDFirma;
            //jpk.TrNLp = dok.GIDLp;
            //jpk.TrNNumer = dok.GIDNumer;
            //jpk.TrNTyp = dok.GIDTyp;
            jpk.Wartosc = "WEW";

            blad = cdn_api.cdn_api.XLDodajKodJPKV7(dokID, jpk);
            if (blad != 0)
            {
                XLKomunikatInfo_20220 kom = new XLKomunikatInfo_20220();
                kom.Wersja = XLWersjaAPI;
                kom.Funkcja = 174;
                kom.Blad = blad;

                cdn_api.cdn_api.XLOpisBledu(kom);
                MessageBox.Show("Błąd jpk dla dokumentu " + Convert.ToString(dataTable.Rows[k][4]) + ":  " + kom.OpisBledu, "Komunikat",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning);


            }*/


            XLZamkniecieDokumentuInfo_20220 dokZam = new XLZamkniecieDokumentuInfo_20220();
            dokZam.Wersja = XLWersjaAPI;
            

            blad=cdn_api.cdn_api.XLZamknijDokument(dokID, dokZam);
            if (blad != 0)
            {
                MessageBox.Show("Błąd zamykania dokumentu " + Convert.ToString(dataTable.Rows[k][4]) + ":  " + blad, "Komunikat",
                           MessageBoxButtons.OK, MessageBoxIcon.Warning);
                

            }


            string update = "update cdn.TraNag set TrN_DokTypJPK = 'WEW' WHERE TrN_GIDTyp = "+dok.GIDTyp+" and TrN_GIDNumer = "+dok.GIDNumer;
            Update(update, connectionString);



        }


        public string ConnectToSQL(string connection_string, string query)
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
        #endregion

        #region funkcje wybranie pliku
        void WybraniePliku()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory(),
                DefaultExt = "xlsx",
                Filter = "Pliki xlsx (*.xlsx)|*.xlsx|Pliki xls (*.xls)|*.xls|Wszystkie pliki (*.*)|*.*",
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tB1.Text = openFileDialog1.FileName;
                //MessageBox.Show(path);
            }
            return;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            WybraniePliku();
        }
        private void TextBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            WybraniePliku();
        }
        #endregion

        #region funkcje pobranie danych z pliku
        bool SprawdzamCzyArkuszNieJestPusty(DataTable _dataTable)
        {
            if (_dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Wybrany arkusz pusty.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                return true;
            }
            return false;
        }


        string znajdzNazwe(string path)
        {
            string[] nazwy = path.Split('\\');
            string nazwa = nazwy[nazwy.Length - 1];
            return nazwa.Substring(0, nazwa.IndexOf(".xls"));
        }

        private void title1_Click(object sender, EventArgs e)
        {

        }

        private void ONE034Form_Load(object sender, EventArgs e)
        {

        }



        public static string Update(string sql, string connection_string)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.UpdateCommand = new SqlCommand(sql, connection);
                    adapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row updated !! ");
                }
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
            return sql;
        }

        /*
        public bool CzyTraNag(short typ)        //zwraca informacje, czy zaznaczony dokument będzie się znajdował w tabeli TraNag
        {
            if (typ == 1529
            || typ == 434
            || typ == 800
            || typ == 1521
            || typ == 1529
            || typ == 1824
            || typ == 1828
            || typ == 1832
            || typ == 1836
            || typ == 2033
            || typ == 2034
            || typ == 2037
            || typ == 2041
            || typ == 2042
            || typ == 2045
            || typ == 2768
            || typ == 2832
            || typ == 2976
            || typ == 2977
            || typ == 3344
            || typ == 3352
            || typ == 4145
            || typ == 7680
            || typ == 7684)
                return true;
            else return false;
        }*/

        #endregion
    }
}