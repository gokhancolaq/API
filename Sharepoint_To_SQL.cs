using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace SharepointAktarim_Ziyaret
{
    class Program
    {


        static void Main(string[] args)
        {

            try
            {
                string connString = @"Data Source=192.168.1.1;Initial Catalog=TESTDB;User ID=admin;Password=admin123";
                SqlConnection connection = new SqlConnection(connString);
                connection.Open();
                SqlCommand verisil = new SqlCommand("TRUNCATE TABLE ZiyaretListesi TRUNCATE TABLE EgitimListesi", connection);
                verisil.ExecuteNonQuery();
                connection.Close();
                connection.Dispose();





                string userName = "admin@microsoft.com";
                string password = "admin123";
                SecureString securePassword = new SecureString();
                foreach (char c in password)
                {
                    securePassword.AppendChar(c);
                };

                string siteUrl_ZY = "https://admin.sharepoint.com/musteri";
                string siteUrl_EG = "https://admin.sharepoint.com/musteri";

                ClientContext clientContext_ZY = new ClientContext(siteUrl_ZY);
                ClientContext clientContext_EG = new ClientContext(siteUrl_EG);

                clientContext_ZY.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                SP.List oList_ZY = clientContext_ZY.Web.Lists.GetByTitle("ZiyaretListesi");

                clientContext_EG.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                SP.List oList_EG = clientContext_ZY.Web.Lists.GetByTitle("EgitimListesi");


                CamlQuery camlQuery_ZY = new CamlQuery();
                ListItemCollection collListItem_ZY = oList_ZY.GetItems(camlQuery_ZY);

                CamlQuery camlQuery_EG = new CamlQuery();
                ListItemCollection collListItem_EG = oList_EG.GetItems(camlQuery_EG);

                clientContext_ZY.Load(
                    collListItem_ZY,
                    items => items.Take(100000).Include(
                    item => item["ZY_TalepEden_AdSoyad"],
                    item => item["ZY_TalepEden_Mail"],
                    item => item["ZY_CariKod"],
                    item => item["ZY_CariUnvan"],
                    item => item["ZY_Cari_il"],
                    item => item["ZY_Sebebi"],
                    item => item["ZY_Not"],
                    item => item["ZY_Baslama_Tarih"],
                    item => item["ZY_Bitis_Tarih"],
                    item => item["ZY_Durum"],
                    item => item["ZY_Gerceklesme"],
                    item => item["ZY_Son_Not"],
                    item => item["ZG_Marka"]
                    ));
                clientContext_ZY.ExecuteQuery();


                clientContext_EG.Load(
                  collListItem_EG,
                  items => items.Take(100000).Include(
                  item => item["Eg_TalepEden_AdSoyad"],
                  item => item["Eg_TalepEden_Mail"],
                  item => item["Eg_CariKod"],
                  item => item["Eg_CariUnvan"],
                  item => item["Eg_Talebi"],
                  item => item["Eg_Not"],
                  item => item["Eg_Baslama_Tarih"],
                  item => item["Eg_Bitis_Tarih"],
                  item => item["Eg_Durum"],
                  item => item["Eg_Tip"],
                  item => item["Eg_Katilimci_Sayisi"],
                  item => item["Eg_Bayi_Fiziki_Puan"],
                  item => item["Onay_Durum"],
                  item => item["Eg_SonNot"],
                  item => item["OdyOrtalama"],
                  item => item["SonPuanOrtalama"],
                  item => item["Eg_Marka"]
                                    ));
                clientContext_EG.ExecuteQuery();


                foreach (ListItem oListItem in collListItem_ZY)
                {
                    string talepedenad, talepedenmail, carikod, cariunvan, il, ziyaretsebebi, not, durum, gerceklesme, sonnot, marka;


                    if (oListItem["ZY_TalepEden_AdSoyad"] == null)
                    {
                        talepedenad = "";
                    }
                    else
                    {
                        talepedenad = oListItem["ZY_TalepEden_AdSoyad"].ToString();

                    }

                    if (oListItem["ZY_TalepEden_Mail"] == null)
                    {
                        talepedenmail = "";
                    }
                    else
                    {
                        talepedenmail = oListItem["ZY_TalepEden_Mail"].ToString();

                    }


                    if (oListItem["ZY_CariKod"] == null)
                    {
                        carikod = "";
                    }
                    else
                    {
                        carikod = oListItem["ZY_CariKod"].ToString();

                    }



                    if (oListItem["ZY_CariUnvan"] == null)
                    {
                        cariunvan = "";
                    }
                    else
                    {
                        cariunvan = oListItem["ZY_CariUnvan"].ToString();

                    }


                    if (oListItem["ZY_Cari_il"] == null)
                    {
                        il = "";
                    }
                    else
                    {
                        il = oListItem["ZY_Cari_il"].ToString();

                    }


                    if (oListItem["ZY_Sebebi"] == null)
                    {
                        ziyaretsebebi = "";
                    }
                    else
                    {
                        ziyaretsebebi = oListItem["ZY_Sebebi"].ToString();

                    }


                    if (oListItem["ZY_Not"] == null)
                    {
                        not = "";
                    }
                    else
                    {
                        not = oListItem["ZY_Not"].ToString();

                    }



                    if (oListItem["ZY_Durum"] == null)
                    {
                        durum = "";
                    }
                    else
                    {
                        durum = oListItem["ZY_Durum"].ToString();

                    }





                    if (oListItem["ZY_Gerceklesme"] == null)
                    {
                        gerceklesme = "";
                    }
                    else
                    {
                        gerceklesme = oListItem["ZY_Gerceklesme"].ToString();

                    }


                    if (oListItem["ZY_Son_Not"] == null)
                    {
                        sonnot = "";
                    }
                    else
                    {
                        sonnot = oListItem["ZY_Son_Not"].ToString();

                    }



                    if (oListItem["ZG_Marka"] == null)
                    {
                        marka = "";
                    }
                    else
                    {
                        marka = oListItem["ZG_Marka"].ToString();

                    }




                    Aktarim_ZY(
                        talepedenad,
                        talepedenmail,
                        carikod,
                        cariunvan,
                        il,
                        ziyaretsebebi,
                        not,
                          Convert.ToDateTime(oListItem["ZY_Baslama_Tarih"].ToString()).AddHours(3),
                        Convert.ToDateTime(oListItem["ZY_Bitis_Tarih"].ToString()).AddHours(3),
                        durum,
                        gerceklesme,
                        sonnot,
                        marka
                        ); ;


                }


                foreach (ListItem oListItem in collListItem_EG)
                {
                    string egtalepedenad, egtalepedenmail, egcarikod, egtip, egkatilimci, egcariunvan, egitimsebebi, egnot, egdurum, egmarka, egfizikpuan, egodyortalama, egsonpuanortalama;


                    if (oListItem["Eg_TalepEden_AdSoyad"] == null)
                    {
                        egtalepedenad = "";
                    }
                    else
                    {
                        egtalepedenad = oListItem["Eg_TalepEden_AdSoyad"].ToString();

                    }

                    if (oListItem["Eg_TalepEden_Mail"] == null)
                    {
                        egtalepedenmail = "";
                    }
                    else
                    {
                        egtalepedenmail = oListItem["Eg_TalepEden_Mail"].ToString();

                    }
                    if (oListItem["Eg_CariKod"] == null)
                    {
                        egcarikod = "";
                    }
                    else
                    {
                        egcarikod = oListItem["Eg_CariKod"].ToString();

                    }
                    if (oListItem["Eg_CariUnvan"] == null)
                    {
                        egcariunvan = "";
                    }
                    else
                    {
                        egcariunvan = oListItem["Eg_CariUnvan"].ToString();

                    }
                    if (oListItem["Eg_Talebi"] == null)
                    {
                        egitimsebebi = "";
                    }
                    else
                    {
                        egitimsebebi = oListItem["Eg_Talebi"].ToString();

                    }
                    if (oListItem["Eg_Not"] == null)
                    {
                        egnot = "";
                    }
                    else
                    {
                        egnot = oListItem["Eg_Not"].ToString();

                    }
                    if (oListItem["Eg_Durum"] == null)
                    {
                        egdurum = "";
                    }
                    else
                    {
                        egdurum = oListItem["Eg_Durum"].ToString();

                    }
                    if (oListItem["Eg_Tip"] == null)
                    {
                        egtip = "";
                    }
                    else
                    {
                        egtip = oListItem["Eg_Tip"].ToString();

                    }
                    if (oListItem["Eg_Katilimci_Sayisi"] == null)
                    {
                        egkatilimci = "";
                    }
                    else
                    {
                        egkatilimci = oListItem["Eg_Katilimci_Sayisi"].ToString();

                    }
                    if (oListItem["Eg_Bayi_Fiziki_Puan"] == null)
                    {
                        egfizikpuan = "";
                    }
                    else
                    {
                        egfizikpuan = oListItem["Eg_Bayi_Fiziki_Puan"].ToString();

                    }
                    if (oListItem["OdyOrtalama"] == null)
                    {
                        egodyortalama = "";
                    }
                    else
                    {
                        egodyortalama = oListItem["OdyOrtalama"].ToString();

                    }
                    if (oListItem["SonPuanOrtalama"] == null)
                    {
                        egsonpuanortalama = "";
                    }
                    else
                    {
                        egsonpuanortalama = oListItem["SonPuanOrtalama"].ToString();

                    }
                    if (oListItem["Eg_Marka"] == null)
                    {
                        egmarka = "";
                    }
                    else
                    {
                        egmarka = oListItem["Eg_Marka"].ToString();

                    }



                    Aktarim_EG(
                        egtalepedenad,
                        egcarikod,
                        egitimsebebi,
                        egnot,
                        Convert.ToDateTime(oListItem["Eg_Baslama_Tarih"].ToString()).AddHours(3),
                        Convert.ToDateTime(oListItem["Eg_Bitis_Tarih"].ToString()).AddHours(3),
                        egdurum,
                        egtip,
                        egkatilimci,
                        egfizikpuan,
                        egodyortalama,
                        egsonpuanortalama,
                        egmarka
                                           
                                              
                        ); ;




                }
            }
            catch
            {
                Environment.Exit(0);
            }





            void Aktarim_ZY(string talepedenad, string talepedenmail, string carikod, string cariunvan, string il, string sebebi, string not, DateTime baslamatarih, DateTime bitistarihi, string durum, string gerceklesme, string sonnot, string marka)
            {


                try
                {
                    string connString2 = @"Data Source=192.168.1.1;Initial Catalog=TESTDB;User ID=admin;Password=admin123";
                    SqlConnection connection2 = new SqlConnection(connString2);
                    DataTable dataTable = new DataTable();
                    string query = "INSERT INTO [dbo].[ZiyaretListesi](ZY_TalepEden_AdSoyad,ZY_TalepEden_Mail,ZY_CariKod,ZY_CariUnvan,ZY_Cari_il,ZY_Sebebi,ZY_Not,ZY_Baslama_Tarih,ZY_Bitis_Tarih,ZY_Durum,ZY_Gerceklesme,ZY_Son_Not,ZG_Marka) VALUES(@ZY_TalepEden_AdSoyad,@ZY_TalepEden_Mail,@ZY_CariKod,@ZY_CariUnvan,@ZY_Cari_il,@ZY_Sebebi,@ZY_Not,@ZY_Baslama_Tarih,@ZY_Bitis_Tarih,@ZY_Durum,@ZY_Gerceklesme,@ZY_Son_Not,@ZG_Marka )";
                    connection2.Open();
                    SqlCommand cmd = new SqlCommand(query, connection2);


                    cmd.Parameters.AddWithValue("@ZY_TalepEden_AdSoyad", talepedenad);
                    cmd.Parameters.AddWithValue("@ZY_TalepEden_Mail", talepedenmail);
                    cmd.Parameters.AddWithValue("@ZY_CariKod", carikod);
                    cmd.Parameters.AddWithValue("@ZY_CariUnvan", cariunvan);
                    cmd.Parameters.AddWithValue("@ZY_Cari_il", il);
                    cmd.Parameters.AddWithValue("@ZY_Sebebi", sebebi);
                    cmd.Parameters.AddWithValue("@ZY_Not", not);
                    cmd.Parameters.AddWithValue("@ZY_Baslama_Tarih", baslamatarih);
                    cmd.Parameters.AddWithValue("@ZY_Bitis_Tarih", bitistarihi);
                    cmd.Parameters.AddWithValue("@ZY_Durum", durum);
                    cmd.Parameters.AddWithValue("@ZY_Gerceklesme", gerceklesme);
                    cmd.Parameters.AddWithValue("@ZY_Son_Not", sonnot);
                    cmd.Parameters.AddWithValue("@ZG_Marka", marka);



                    cmd.ExecuteNonQuery();

                    connection2.Close();
                    connection2.Dispose();
                }

                catch
                {

                    Environment.Exit(0);
                }


            }

            void Aktarim_EG(string egtalepedenad,  string egcarikod,  string egitimsebebi, string egnot, DateTime egbaslamatarih, DateTime egbitistarihi, string egdurum, string egtip, string egkatilimci, string egfizikpuan, string egodyortalama, string egsonpuanortalama, string egmarka)
                {

                    try
                    {
                        string connString3 = @"Data Source=192.168.1.1;Initial Catalog=TESTDB;User ID=admin;Password=admin123";
                        SqlConnection connection3 = new SqlConnection(connString3);
                        DataTable dataTableEG = new DataTable();
                        string query = "INSERT INTO [dbo].[EgitimListesi]" +
                            "(Eg_TalepEden_AdSoyad, Eg_CariKod, Eg_Talebi, Eg_Not, Eg_Baslama_Tarih, Eg_Bitis_Tarih, Eg_Durum, Eg_Tip, Eg_Katilimci_Sayisi, Eg_Bayi_Fiziki_Puan,   OdyOrtalama, SonPuanOrtalama,Eg_Marka)" +
                            "VALUES" +
                            "(@Eg_TalepEden_AdSoyad, @Eg_CariKod, @Eg_Talebi, @Eg_Not, @Eg_Baslama_Tarih, @Eg_Bitis_Tarih, @Eg_Durum, @Eg_Tip, @Eg_Katilimci_Sayisi, @Eg_Bayi_Fiziki_Puan,   @OdyOrtalama, @SonPuanOrtalama,@Eg_Marka)";
                        connection3.Open();
                        SqlCommand cmd = new SqlCommand(query, connection3);
                        cmd.Parameters.AddWithValue("@Eg_TalepEden_AdSoyad", egtalepedenad);
                        cmd.Parameters.AddWithValue("@Eg_CariKod", egcarikod);
                        cmd.Parameters.AddWithValue("@Eg_Talebi", egitimsebebi);
                        cmd.Parameters.AddWithValue("@Eg_Not", egnot);
                        cmd.Parameters.AddWithValue("@Eg_Baslama_Tarih", egbaslamatarih);
                        cmd.Parameters.AddWithValue("@Eg_Bitis_Tarih", egbitistarihi);
                        cmd.Parameters.AddWithValue("@Eg_Durum", egdurum);
                        cmd.Parameters.AddWithValue("@Eg_Tip", egtip);
                        cmd.Parameters.AddWithValue("@Eg_Katilimci_Sayisi", egkatilimci);
                        cmd.Parameters.AddWithValue("@Eg_Bayi_Fiziki_Puan", egfizikpuan);
                        cmd.Parameters.AddWithValue("@OdyOrtalama", egodyortalama);
                        cmd.Parameters.AddWithValue("@SonPuanOrtalama", egsonpuanortalama);
                        cmd.Parameters.AddWithValue("@Eg_Marka", egmarka);


                        cmd.ExecuteNonQuery();

                        connection3.Close();
                        connection3.Dispose();
                    }

                    catch
                    {

                        Environment.Exit(0);
                    }


                }

        }
    }
}

