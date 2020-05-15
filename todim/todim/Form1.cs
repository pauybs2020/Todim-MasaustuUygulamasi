using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections; //arraylist kullanımı için ekledim
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.WinFormsUtilities;
using System.IO;
using System.Reflection;
using OfficeOpenXml;

namespace todim
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY"); //import için

            InitializeComponent();
            gridTasarimSirasiz(dataGridViewKararMat);
            gridTasarim(dataGridViewAgirlik);
            gridTasarim(dataGridViewNormalize);
            gridTasarim(dataGridViewBaskinlikSkoru);
            gridTasarim(dataGridViewGenelSkor);
            gridTasarim(dataGridViewKarsilastirmaMat);
            gridTasarim(dataGridViewC);
            gridTasarim(dataGridViewWVektörü);
            gridTasarim(dataGridViewAyrintiKmat);
            gridTasarim(dataGridViewDVektör);
            gridTasarim(dataGridViewKriterAgirliklari);
            gridTasarim(dataGridViewSonucKararMat);
            gridTasarim(dataGridViewSonucNormalizeMat);
            gridTasarim(dataGridViewSonucKriterAgirlik);
            gridTasarim(dataGridViewSonucGoreliAgirlik);
            gridTasarim(dataGridViewSonucGenelBaskinlik);
            gridTasarim(dataGridViewSonucKarsilastirmaMat);

        }
        public static ArrayList kriterler = new ArrayList();
        public static ArrayList alternatifler = new ArrayList();
        public static ArrayList faydaMaliyet = new ArrayList();
        public static ArrayList agirliklar = new ArrayList();
        public static ArrayList baskinlikSkorlari = new ArrayList();
        public static ArrayList tBaskinlikSkorlari = new ArrayList();
        public static ArrayList goreliAgirliklar = new ArrayList(); //Kriterlerin referans kriterine olan göreli ağırlıklarını(wjr) bu listeye atadım
        public static ArrayList maxList = new ArrayList();
        public static ArrayList minList = new ArrayList();
        public static ArrayList genelBaskinlikSkorlari = new ArrayList();
        public static ArrayList paydaListesi = new ArrayList(); //yüzde önem dağılımlarını hesaplamak için gereken sutun toplamlarını tutar       
        public static ArrayList kararMatSutunToplam = new ArrayList();
        double wjr, wjr1; //Bir kriterin referans kriterine olan göreli ağırlığı
        double max, min, gAgirlikTop, wjrToplam, rijFark, sonuc, sonuc1, satirSkorToplam, lamda, CI, CR;
        double baskinlikSkoru, tbaskinlikSkoru; // Ai alternatifinin diğer alternatiflere baskınlık skoru          
        double wr; //referans kriterinin ağırlığı (çalışmada en yüksek ağırlık değeri referans kriteri olarak alınmış)
        double maxBaskinlikSkoru, minBaskinlikSkoru, genelBaskinlikSkoru, baskinlikSkorum, bskor, maxGenelBaskinlikSkor, gSkor, agirlikToplam;
        string enİyiAlternatif, alternatifBskor, girilenAlternatif, calismaIsmi, tiklanma, yontem, secilenNormalizeYontemi;
        int rbtnDiziboyut, rbtnDizi1boyut = 1;
        int x, y, rbtn, rbtn1, satir = 0;
        RadioButton[] radioButton;
        RadioButton[] radioButton1;
        int duzenleIndex; //kriter ve alternatiflerin düzenlenmesi için listboxdaki seçili index i tutan değişken
        int baskinlikSkorButon; //tüm baskınlık skorlarının excel e aktarılmasında hangi butona tıklanıldığını anlamak için tanımladığım değişken
        //datagridView tasarım kodları
        DataGridView dgvTumBSkor;
        DataGridView[] dgvTumBSkorDizi;
        double θ = 1;
        string agrYonSakla, normalizeYonSakla, sutunHarfi;
        string[] excelSutun = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ" };

        public void verileriCagir()
        {

            //string[] excelSutun = { "A","B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ" };


            int say = 0;
            int sayKriterEkle = 0;


            tumunuTemizle(); //var
            tiklanma = "esli";//var
            kararMatGorunurlukAyarlari();//var


            OpenFileDialog OFD = new OpenFileDialog()
            {
                Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                Title = "Excel Dosyası Seçiniz..",
                RestoreDirectory = true,
            };

            if (OFD.ShowDialog() == DialogResult.OK)

            {
                string DosyaYolu = OFD.FileName;// dosya yolu
                string DosyaAdi = OFD.SafeFileName; // dosya adı

                OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DosyaYolu + ";   Extended Properties =\"Excel 12.0;HDR=Yes;IMEX=0\"");
                baglanti.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
                DataTable DTexcel = new DataTable();
                da.Fill(DTexcel);
                dataGridViewKararMat.DataSource = DTexcel;

                kararMatListeDoldurmaEskiCalisma();
                baglanti.Close();



                faydaMaliyet.Clear();
                OleDbConnection baglanti2 = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DosyaYolu + ";   Extended Properties =\"Excel 12.0;HDR=No;\"");
                baglanti2.Open();
                string sql2 = "select * from [Sayfa2$A1:A" + kriterler.Count + 1 + "] ";
                OleDbCommand veri2 = new OleDbCommand(sql2, baglanti2); OleDbDataReader dr = null;
                dr = veri2.ExecuteReader();


                for (int i = 0; i < kriterler.Count; i++)
                {
                    dr.Read();
                    faydaMaliyet.Add(dr[0].ToString());
                }

                string sql3 = "select * from [Sayfa3$A1:A2] ";
                OleDbCommand veri3 = new OleDbCommand(sql3, baglanti2); OleDbDataReader dr3 = null;
                dr3 = veri3.ExecuteReader();
                dr3.Read();
                secilenNormalizeYontemi = dr3[0].ToString();
                normalizeYonSakla = dr3[0].ToString();

                string sql4 = "select * from [Sayfa3$A2:A3] ";
                OleDbCommand veri4 = new OleDbCommand(sql4, baglanti2); OleDbDataReader dr4 = null;
                dr4 = veri4.ExecuteReader();
                dr4.Read();
                yontem = dr4[0].ToString();
                agrYonSakla = dr4[0].ToString();


                agirliklar.Clear();
                string sql5 = "select * from [Sayfa4$A1:A" + kriterler.Count + 1 + "] ";
                //string sql5 = "select * from [Sayfa4$]";
                OleDbCommand veri5 = new OleDbCommand(sql5, baglanti2); OleDbDataReader dr5 = null;
                dr5 = veri5.ExecuteReader();
                for (int i = 0; i < kriterler.Count; i++)
                {
                    dr5.Read();
                    agirliklar.Add(dr5[0].ToString());
                }
                sutunHarfi = "";
                if (yontem == "ahp")
                {
                    baglanti.Open();

                    //ikili karşılaştırma matrisi için
                    say = kriterler.Count + 1;
                    sutunHarfi = (excelSutun[kriterler.Count + 1]).ToString();
                    OleDbDataAdapter da2 = new OleDbDataAdapter("SELECT * FROM [Sayfa5$A1:" + sutunHarfi + say + "] ", baglanti);
                    //OleDbDataAdapter da2 = new OleDbDataAdapter("SELECT * FROM [Sayfa5$]", baglanti);
                    DataTable DTexcel2 = new DataTable();
                    da2.Fill(DTexcel2);
                    dataGridViewAyrintiKmat.DataSource = DTexcel2;
                    dataGridViewKarsilastirmaMat.DataSource = DTexcel2;
                    dataGridViewKarsilastirmaMat.Columns[0].HeaderText = "";
                    dataGridViewAyrintiKmat.Columns[0].HeaderText = "";


                    //c matrisi için
                    say += 2;
                    sayKriterEkle = say + kriterler.Count;
                    OleDbDataAdapter da3 = new OleDbDataAdapter("SELECT * FROM [Sayfa5$A" + say + ":" + sutunHarfi + sayKriterEkle + "] ", baglanti);
                    DataTable DTexcel3 = new DataTable();
                    da3.Fill(DTexcel3);
                    dataGridViewC.DataSource = DTexcel3;
                    dataGridViewC.Columns[0].HeaderText = "";


                    //ağırlık değerleri için
                    say += kriterler.Count + 2;
                    sayKriterEkle = say + 1;
                    OleDbDataAdapter da4 = new OleDbDataAdapter("SELECT * FROM [Sayfa5$A" + say + ":" + sutunHarfi + sayKriterEkle + "] ", baglanti);
                    DataTable DTexcel4 = new DataTable();
                    da4.Fill(DTexcel4);
                    dataGridViewWVektörü.DataSource = DTexcel4;


                    //d vektörü için
                    say += 3;
                    sayKriterEkle = say + 1;
                    OleDbDataAdapter da5 = new OleDbDataAdapter("SELECT * FROM [Sayfa5$A" + say + ":" + sutunHarfi + sayKriterEkle + "] ", baglanti);
                    DataTable DTexcel5 = new DataTable();
                    da5.Fill(DTexcel5);
                    dataGridViewDVektör.DataSource = DTexcel5;

                    // CI Değeri için
                    say += 3;
                    sayKriterEkle = say + 1;
                    string sql6 = "SELECT * FROM [Sayfa5$A" + say + ":" + sutunHarfi + sayKriterEkle + "] ";
                    OleDbCommand veri6 = new OleDbCommand(sql6, baglanti2); OleDbDataReader dr6 = null;
                    dr6 = veri6.ExecuteReader();
                    dr6.Read();
                    //CI = Convert.ToDouble(dr6[0].ToString());
                    lblCI.Text = dr6[0].ToString();

                    //RI değeri için
                    say += 2;
                    sayKriterEkle = say + 1;
                    string sql7 = "SELECT * FROM [Sayfa5$A" + say + ":" + sutunHarfi + sayKriterEkle + "] ";
                    OleDbCommand veri7 = new OleDbCommand(sql7, baglanti2); OleDbDataReader dr7 = null;
                    dr7 = veri7.ExecuteReader();
                    dr7.Read();
                    lblRI.Text = dr7[0].ToString();

                    //tutarlılık oranı için
                    say += 2;
                    sayKriterEkle = say + 1;
                    string sql8 = "SELECT * FROM [Sayfa5$A" + say + ":" + sutunHarfi + sayKriterEkle + "] ";
                    OleDbCommand veri8 = new OleDbCommand(sql8, baglanti2); OleDbDataReader dr8 = null;
                    dr8 = veri8.ExecuteReader();
                    dr8.Read();
                    lblTutarlilikOrani.Text = dr8[0].ToString();
                    CR = Convert.ToDouble(Convert.ToDouble(lblCI.Text) / Convert.ToDouble(lblRI.Text));
                    lblTutOrani.Text = CR.ToString();


                    ahpTasarim(); //KARŞILAŞTIRMA MATRİSİ OLUŞTURMA EKRANI

                }
                //TETA DEĞERİNİ KAYDETMEK İÇİN 
                string sql9 = "select * from [Sayfa3$A3:A4] ";
                OleDbCommand veri9 = new OleDbCommand(sql9, baglanti2); OleDbDataReader dr9 = null;
                dr9 = veri9.ExecuteReader();
                dr9.Read();
                θ = Convert.ToDouble(dr9[0]);

                baglanti2.Close();



                kararMatRenklendir();
                boyutAyarlama();
                tabControl1.SelectedTab = tabPageKararMatrisi;
                gridTasarimSirasiz(dataGridViewKararMat);
                normalizeMatCerceve();
                kararMatrisiToolStripMenuItem1.Visible = true;


                //karar matrisini doldurduktan kaydettiğim normalize yöntemine göre normalize edilmiş karar matrisini doldurmam lazım
                normalizeMatCerceve();
                if (secilenNormalizeYontemi == "1")
                {
                    maxMin(); //her sutundaki max ve min değerleri bulup ilgili listelere atayan metod
                              //normalizasyon
                    double max1, min1;
                    for (int j = 1; j < kriterler.Count + 1; j++)
                    {
                        if (faydaMaliyet[j - 1].ToString() == "Fayda")
                        {
                            max1 = Convert.ToDouble(maxList[j - 1]);
                            min1 = Convert.ToDouble(minList[j - 1]);

                            for (int i = 0; i < alternatifler.Count; i++)
                            {
                                if (max1 - min1 == 0)
                                {
                                    dataGridViewNormalize.Rows[i].Cells[j].Value = 0;
                                }
                                else
                                {
                                    dataGridViewNormalize.Rows[i].Cells[j].Value = (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value) - min1) / (max1 - min1);

                                }
                            }
                        }
                        else /*if (faydaMaliyet[j - 1].ToString() == "Maliyet")*/
                        {
                            max1 = Convert.ToDouble(maxList[j - 1]);
                            min1 = Convert.ToDouble(minList[j - 1]);
                            for (int i = 0; i < alternatifler.Count; i++)
                            {
                                if (max1 - min1 == 0)
                                {
                                    dataGridViewNormalize.Rows[i].Cells[j].Value = 0;
                                }
                                else
                                {
                                    dataGridViewNormalize.Rows[i].Cells[j].Value = (max1 - (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value))) / (max1 - min1);
                                }

                            }

                        }
                    }


                }

                else if (secilenNormalizeYontemi == "2")
                {


                    for (int j = 1; j < kriterler.Count + 1; j++)
                    {
                        if (faydaMaliyet[j - 1].ToString() == "Fayda")
                        {
                            double sutunToplam = 0;
                            for (int i = 0; i < alternatifler.Count; i++)
                            {
                                sutunToplam += Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value);
                            }
                            kararMatSutunToplam.Add(sutunToplam);
                        }
                        else
                        {
                            double sutunBireBolTopla = 0;
                            for (int i = 0; i < alternatifler.Count; i++)
                            {
                                sutunBireBolTopla += Convert.ToDouble(1 / (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value)));
                            }
                            kararMatSutunToplam.Add(sutunBireBolTopla);
                        }
                    }

                    //normalizasyon

                    for (int j = 1; j < kriterler.Count + 1; j++)
                    {
                        for (int i = 0; i < alternatifler.Count; i++)
                        {
                            double sayi = Convert.ToDouble(kararMatSutunToplam[j - 1]);
                            if (faydaMaliyet[j - 1].ToString() == "Fayda")
                            {
                                if (Convert.ToDouble(kararMatSutunToplam[j - 1]) == 0)
                                {
                                    dataGridViewNormalize.Rows[i].Cells[j].Value = 0;
                                }
                                else
                                {
                                    dataGridViewNormalize.Rows[i].Cells[j].Value = (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value)) / sayi;
                                }
                            }
                            else /*if (faydaMaliyet[j - 1].ToString() == "Maliyet")*/
                            {
                                if (Convert.ToDouble(kararMatSutunToplam[j - 1]) == 0)
                                {
                                    dataGridViewNormalize.Rows[i].Cells[j].Value = 0;
                                }
                                else
                                {
                                    dataGridViewNormalize.Rows[i].Cells[j].Value = (1 / (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value))) / sayi;
                                }
                            }
                        }
                    }

                }


                //kaydedilen ağırlık yöntemine göre ağırlık matrislerini doldurma

                if (yontem == "manuel")
                {
                    manuelAgirlikMatrisi();
                    //matrisin içini excelden çektiğim ve listeye attığım ağırlık değerleriyle doldurmam lazım

                    for (int j = 1; j < kriterler.Count + 1; j++)
                    {
                        dataGridViewAgirlik.Rows[0].Cells[j].Value = Convert.ToDouble(agirliklar[j - 1]);
                    }
                    //gizli olan panelleri görünür yaptım
                    panel12.Visible = true;

                    pnlYontemSec.Visible = true;

                }

                if (yontem == "ahp")
                {

                    pnlYontemSec.Visible = true;
                    btnIkiliKMatEAc.Visible = true;
                    btnIkiliKMatEAktar.Visible = true;
                    btnKriterAgirlikKaydet.Visible = true;
                    btnAhpAyrintiCozGoster.Visible = true;
                    dataGridViewKriterAgirliklari.Rows.Clear();
                    panel82.Visible = true;
                    panel81.Visible = true;
                    panel83.Visible = true;
                    dgvKriterAgirlikDoldur();
                    panel80.Visible = true;
                    label24.Visible = true;
                    lblTutOrani.Visible = true;


                    if (Convert.ToDouble(lblTutOrani.Text) >= 0.10)
                    {
                        lblUyari.Visible = true;
                        panel52.Visible = true;
                        btnKriterAgirlikKaydet.Text = "Bu şekilde devam et";
                    }
                }


                goreliAgirliklarim();
                maxGenelBakinlik();
                genelBaskinlikSkoruMatrisi();

                //SONUÇLAR
                sonucKararMatDoldur();
                sonucNormalizeMatDoldur();
                sonucKriterAgirlikDoldur();
                sonucGoreliAgirlikDoldur();
                sonucGenelBaskinlikDoldur();
                if (yontem == "ahp")
                {
                    sonucKarsilastirmaMatDoldur();
                    AHPToolStripMenuItem.Visible = true;
                    karsilastirmaMatrisiOlusturmaToolStripMenuItem.Visible = true;
                    agirlikDegerleriToolStripMenuItem.Visible = true;
                    ahpAyrintiliCozumToolStripMenuItem.Visible = true;
                }

                //tüm kısmi baskınlık skoru görüntüleme
                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel3.Controls.Clear();
                tumKismiBaskinlikSkorGoruntule();



                baslangicToolStripMenuItem.Visible = true;
                kararMatrisiOlusturmaToolStripMenuItem.Visible = true;
                kararMatrisiToolStripMenuItem1.Visible = true;
                normalizeEdilmisKararMatrisiToolStripMenuItem.Visible = true;
                kriterAgirligiBelirlemeToolStripMenuItem.Visible = true;
                agirlikYontemSecToolStripMenuItem.Visible = true;
                sonuçlarToolStripMenuItem.Visible = true;
                genelBaskinlikSkorlariToolStripMenuItem.Visible = true;
                kismiBaskinlikSkorlariToolStripMenuItem.Visible = true;
                kismiBaskinlikSkoruAramaToolStripMenuItem.Visible = true;
                ayrintiliCozumToolStripMenuItem.Visible = true;



            }
            else
            {
                MessageBox.Show("Dosya seçilmedi!", "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


        }
        public void matrisleriKaydet()
        {
            //try
            //{
            //    ExcelPackage package = new ExcelPackage();
            //    package.Workbook.Worksheets.Add("Sayfa1");
            //    OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            //    int satir = 1;
            //    var columns = dataGridViewSonucKararMat.Columns;
            //    for (int i = 0; i < columns.Count; i++)
            //    {
            //        worksheet.Cells[satir, i + 1].Value = columns[i].HeaderText;
            //    }

            //    satir++;
            //    var rows = dataGridViewSonucKararMat.Rows;
            //    for (int i = 0; i < rows.Count; i++)
            //    {
            //        if (rows[i].Cells[0] != null)
            //        {
            //            for (int j = 0; j < rows[i].Cells.Count; j++)
            //            {
            //                worksheet.Cells[satir, j + 1].Value = rows[i].Cells[j].Value;
            //            }
            //            satir++;
            //        }
            //    }


            //    package.Workbook.Worksheets.Add("Sayfa2");
            //    OfficeOpenXml.ExcelWorksheet worksheet2 = package.Workbook.Worksheets.FirstOrDefault();
            //    satir++;


            //    for (int i = 0; i < dataGridViewSonucKriterAgirlik.Columns.Count; i++)
            //    {
            //        worksheet2.Cells[satir, i + 1].Value = dataGridViewSonucKriterAgirlik.Columns[i].HeaderText;
            //    }

            //    satir++;

            //    for (int i = 0; i < dataGridViewSonucKriterAgirlik.Rows.Count; i++)
            //    {
            //        if (rows[i].Cells[0] != null)
            //        {
            //            for (int j = 0; j < dataGridViewSonucKriterAgirlik.Rows[i].Cells.Count; j++)
            //            {
            //                worksheet2.Cells[satir, j + 1].Value = dataGridViewSonucKriterAgirlik.Rows[i].Cells[j].Value;
            //            }
            //            satir++;
            //        }
            //    }


            //    SaveFileDialog saveFileDialog = new SaveFileDialog();
            //    saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
            //    saveFileDialog.ShowDialog();

            //    Stream stream = saveFileDialog.OpenFile();
            //    package.SaveAs(stream);

            //    stream.Close();

            //    MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}


            //KLASÖR OLUŞTURMA
            string klasorAd = @"c:\TODİM-Uygulamalar";

            //klasör yoksa oluşturuyorum
            if (System.IO.Directory.Exists(klasorAd) == false)
            {
                System.IO.Directory.CreateDirectory(klasorAd);
            }
            //klasör varsa oraya dosyayı kaydediyorum
            if (System.IO.Directory.Exists(klasorAd) == true)
            {

                //MATRİSLERİ KAYDETME

                try
                {


                    var saveFileDialog = new SaveFileDialog();
                    saveFileDialog.InitialDirectory = @"c:\TODİM-Uygulamalar"; // her zaman bu dosya konumunu açması için
                    saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
                    saveFileDialog.FilterIndex = 3;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        var workbook = new ExcelFile();
                        var worksheet = workbook.Worksheets.Add("Sayfa1");
                        DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewSonucKararMat, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

                        var worksheet2 = workbook.Worksheets.Add("Sayfa2");
                        int i = 0;
                        foreach (var item in faydaMaliyet)
                        {

                            worksheet2.Cells[i, 0].Value = item.ToString();
                            i++;
                        }


                        var worksheet3 = workbook.Worksheets.Add("Sayfa3");
                        worksheet3.Cells[0, 0].Value = secilenNormalizeYontemi.ToString();
                        worksheet3.Cells[1, 0].Value = yontem.ToString();
                        worksheet3.Cells[2, 0].Value = θ.ToString();
                        var worksheet4 = workbook.Worksheets.Add("Sayfa4");
                        int j = 0;
                        foreach (var item in agirliklar)
                        {

                            worksheet4.Cells[j, 0].Value = item.ToString();
                            j++;
                        }
                        if (yontem == "ahp")
                        {

                            var worksheet5 = workbook.Worksheets.Add("Sayfa5");
                            int say = 0;

                            DataGridViewConverter.ImportFromDataGridView(worksheet5, this.dataGridViewAyrintiKmat, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                            say += dataGridViewAyrintiKmat.Rows.Count + 2;

                            DataGridViewConverter.ImportFromDataGridView(worksheet5, this.dataGridViewC, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                            say += dataGridViewC.Rows.Count + 2;

                            DataGridViewConverter.ImportFromDataGridView(worksheet5, this.dataGridViewWVektörü, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                            say += dataGridViewWVektörü.Rows.Count + 2;

                            DataGridViewConverter.ImportFromDataGridView(worksheet5, this.dataGridViewDVektör, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                            say += dataGridViewDVektör.Rows.Count + 2;

                            worksheet5.Cells[say, 0].Value = lblCI.Text;
                            say += 2;
                            worksheet5.Cells[say, 0].Value = lblRI.Text;
                            say += 2;
                            worksheet5.Cells[say, 0].Value = lblTutOrani.Text;
                        }

                        workbook.Save(saveFileDialog.FileName);

                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show("Excel'e aktarma işlemi başarısız!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }



        }
        public void agirliklariExceldenAl()
        {
            try
            {
                OpenFileDialog OFD = new OpenFileDialog()
                {
                    Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                    Title = "Excel Dosyası Seçiniz..",
                    RestoreDirectory = true,

                };

                if (OFD.ShowDialog() == DialogResult.OK)
                {

                    string DosyaYolu = OFD.FileName;// dosya yolu
                    string DosyaAdi = OFD.SafeFileName; // dosya adı

                    OleDbConnection baglanti2 = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DosyaYolu + ";   Extended Properties =\"Excel 12.0;HDR=No;\"");
                    baglanti2.Open();
                    sutunHarfi = (excelSutun[kriterler.Count]).ToString();
                    string sql2 = "select * from [Sayfa1$A1:A" + kriterler.Count + 1 + "]";
                    OleDbCommand veri2 = new OleDbCommand(sql2, baglanti2); OleDbDataReader dr = null;
                    dr = veri2.ExecuteReader();


                    for (int i = 0; i < kriterler.Count; i++)

                    {
                        dr.Read();
                        agirliklar.Add(dr[0].ToString());
                    }


                    for (int i = 1; i < kriterler.Count + 1; i++)
                    {
                        dataGridViewAgirlik.Rows[0].Cells[i].Value = Convert.ToDouble(agirliklar[i - 1]);
                    }
                    agirliklar.Clear();


                    //OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DosyaYolu + ";   Extended Properties =\"Excel 12.0;HDR=Yes;IMEX=0\"");
                    //OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
                    //DataTable DTexcel = new DataTable();
                    //da.Fill(DTexcel);
                    //dataGridViewAgirlik.DataSource = DTexcel;
                    //baglanti.Close();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ağırlık getirme işlemi başarısız" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void kararMatExcelYukle()
        {
            OpenFileDialog OFD = new OpenFileDialog()
            {
                Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                // open file dialog açıldığında sadece excel dosyalarınu görecek
                Title = "Excel Dosyası Seçiniz..",
                // open file dialog penceresinin başlığı
                RestoreDirectory = true,
                // en son açtığı klasörü gösterir. Örn en son excel dosyasını D://Exceller adlı
                // bir klasörden çekmiş olsun. Bir sonraki open file dialog açıldığında yine aynı 
                // klasörü gösterecektir.
            };
            // bu da bir kullanım şeklidir. Aslında bu şekilde kullanmak daha faydalıdır. 
            // bir çok intelligence aracı bu şekilde kullanılmasını tavsiye ediyor.
            if (OFD.ShowDialog() == DialogResult.OK)
            // perncere açıldığında dosya seçildi ise yapılacak. Bunu yazmazsak dosya seçmeden 
            // kapandığında program kırılacaktır.
            {
                string DosyaYolu = OFD.FileName;// dosya yolu
                string DosyaAdi = OFD.SafeFileName; // dosya adı

                //OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DosyaYolu 
                //    + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DosyaYolu + ";   Extended Properties =\"Excel 12.0;HDR=Yes;IMEX=0\"");

                // excel dosyasına access db gibi bağlanıyoruz.

                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
                // burada FROM dan sonra sayfa1$ kısmı önemlidir.sayfa adı faklı ise örn
                // sheet ise program hata verecektir.
                // NOT: Excel dosyanızın ilk satır başlık olsun. Yani sistem öyle algıladığından 
                // ilk satırdaki bilgileri başlık olarak tanımlayıp almıyor. Ne yazarsanız yazın
                // sorun teşkil etmiyor. Tabi db için özel olan karakterleri kullanmayın.
                DataTable DTexcel = new DataTable();
                da.Fill(DTexcel);
                // select sorgusu ile okunan verileri datatable'ye aktarıyoruz.

                dataGridViewKararMat.DataSource = DTexcel;
                // datatable'ı da gridcontrol'ün datasource'una atıyoruz.

                baglanti.Close();

            }
        }
        ToolTip bilgiMesaji(string baslik, string aciklama, Control nesne)
        {
            ToolTip bilgi = new ToolTip();
            bilgi.Active = true; //görünürlüğü
            bilgi.ToolTipTitle = baslik; //mesaj başlığı
            bilgi.ToolTipIcon = ToolTipIcon.Info; //ikon 
            bilgi.UseFading = true; //silik olarak kaybolup yüklenme
            bilgi.UseAnimation = true;
            bilgi.IsBalloon = true;
            bilgi.ShowAlways = true; //her zaman göster
            bilgi.AutoPopDelay = 2500; //mesajın açık kalma süresi
            bilgi.ReshowDelay = 2000; //mouse çekildikten kaç ms sonra kaybolacağı
            bilgi.InitialDelay = 700; //mesajın açılma süresi
            bilgi.BackColor = Color.White;
            bilgi.ForeColor = Color.DarkBlue;
            bilgi.SetToolTip(nesne, aciklama); //hangi kontrolde görüneceği


            return bilgi;
        }
        ToolTip bilgiMesajiRadioButton(/*string baslik,*/ string aciklama, Control nesne)
        {
            ToolTip bilgi = new ToolTip();
            bilgi.Active = true; //görünürlüğü
            //bilgi.ToolTipTitle = baslik; //mesaj başlığı
            //bilgi.ToolTipIcon = ToolTipIcon.Info; //ikon 
            bilgi.UseFading = true; //silik olarak kaybolup yüklenme
            bilgi.UseAnimation = true;
            bilgi.IsBalloon = false;
            bilgi.ShowAlways = true; //her zaman göster
            bilgi.AutoPopDelay = 2500; //mesajın açık kalma süresi
            bilgi.ReshowDelay = 500; //mouse çekildikten kaç ms sonra kaybolacağı
            bilgi.InitialDelay = 100; //mesajın açılma süresi
            bilgi.BackColor = Color.White;
            bilgi.ForeColor = Color.DarkBlue;
            bilgi.SetToolTip(nesne, aciklama); //hangi kontrolde görüneceği


            return bilgi;
        }
        public void gridTasarim(DataGridView datagridview)
        {
            datagridview.RowHeadersVisible = false;  //ilk sutunu gizleme
            datagridview.BorderStyle = BorderStyle.None;
            datagridview.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSlateGray;//varsayılan arka plan rengi verme
            datagridview.DefaultCellStyle.SelectionBackColor = Color.Silver; //seçilen hücrenin arkaplan rengini belirleme
            datagridview.DefaultCellStyle.SelectionForeColor = Color.White;  //seçilen hücrenin yazı rengini belirleme
            datagridview.EnableHeadersVisualStyles = false; //başlık özelliğini değiştirme
            datagridview.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; //başlık çizgilerini ayarlama
            datagridview.ColumnHeadersDefaultCellStyle.BackColor = Color.Gray; //başlık arkaplan rengini belirleme
            datagridview.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;  //başlık yazı rengini belirleme
            datagridview.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;                                                                   // datagridview.SelectionMode = DataGridViewSelectionMode.FullRowSelect; //satırı tamamen seçmeyi sağlama
            datagridview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;  //her hangi bir sutunun genişliğini o sutunda yer alan en  uzun yazının genişliğine göre ayarlama
            datagridview.AllowUserToAddRows = false;  //ilk sutunu gizleme
            datagridview.AllowUserToOrderColumns = true;

        }
        public void gridTasarimSirasiz(DataGridView datagridview)
        {
            datagridview.RowHeadersVisible = false;  //ilk sutunu gizleme         
            datagridview.BorderStyle = BorderStyle.None;
            datagridview.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSlateGray;//varsayılan arka plan rengi verme
            datagridview.DefaultCellStyle.SelectionBackColor = Color.Silver; //seçilen hücrenin arkaplan rengini belirleme
            datagridview.DefaultCellStyle.SelectionForeColor = Color.White;  //seçilen hücrenin yazı rengini belirleme
            datagridview.EnableHeadersVisualStyles = false; //başlık özelliğini değiştirme
            datagridview.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single; //başlık çizgilerini ayarlama
            datagridview.ColumnHeadersDefaultCellStyle.BackColor = Color.Gray; //başlık arkaplan rengini belirleme
            datagridview.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;  //başlık yazı rengini belirleme
            datagridview.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;                                                                   // datagridview.SelectionMode = DataGridViewSelectionMode.FullRowSelect; //satırı tamamen seçmeyi sağlama
            datagridview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;  //her hangi bir sutunun genişliğini o sutunda yer alan en  uzun yazının genişliğine göre ayarlama
            datagridview.AllowUserToAddRows = false;
            datagridview.AllowUserToOrderColumns = false;
        }
        public void gridTasarim2(DataGridView datagridview)
        {
            datagridview.RowHeadersVisible = false;  //ilk sutunu gizleme
            datagridview.BorderStyle = BorderStyle.None;
            datagridview.AlternatingRowsDefaultCellStyle.BackColor = Color.White;//varsayılan arka plan rengi verme
            datagridview.DefaultCellStyle.SelectionBackColor = Color.LightGray; //seçilen hücrenin arkaplan rengini belirleme
            datagridview.DefaultCellStyle.SelectionForeColor = Color.White;  //seçilen hücrenin yazı rengini belirleme
            datagridview.EnableHeadersVisualStyles = false; //başlık özelliğini değiştirme
            datagridview.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; //başlık çizgilerini ayarlama
            datagridview.ColumnHeadersDefaultCellStyle.BackColor = Color.RosyBrown; //başlık arkaplan rengini belirleme
            datagridview.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;  //başlık yazı rengini belirleme                                                                               // datagridview.SelectionMode = DataGridViewSelectionMode.FullRowSelect; //satırı tamamen seçmeyi sağlama
            datagridview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;  //her hangi bir sutunun genişliğini o sutunda yer alan en  uzun yazının genişliğine göre ayarlama
            datagridview.AllowUserToAddRows = false;  //ilk sutunu gizleme
            datagridview.AllowUserToOrderColumns = true;
        }
        public void gridTasarim3(DataGridView datagridview)
        {
            datagridview.RowHeadersVisible = false;  //ilk sutunu gizleme
            datagridview.BorderStyle = BorderStyle.None;
            datagridview.AlternatingRowsDefaultCellStyle.BackColor = Color.White;//varsayılan arka plan rengi verme
            datagridview.DefaultCellStyle.SelectionBackColor = Color.LightGray; //seçilen hücrenin arkaplan rengini belirleme
            datagridview.DefaultCellStyle.SelectionForeColor = Color.White;  //seçilen hücrenin yazı rengini belirleme
            datagridview.EnableHeadersVisualStyles = false; //başlık özelliğini değiştirme
            datagridview.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; //başlık çizgilerini ayarlama
            datagridview.ColumnHeadersDefaultCellStyle.BackColor = Color.CadetBlue; //başlık arkaplan rengini belirleme
            datagridview.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;  //başlık yazı rengini belirleme                                                                                // datagridview.SelectionMode = DataGridViewSelectionMode.FullRowSelect; //satırı tamamen seçmeyi sağlama
            datagridview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;  //her hangi bir sutunun genişliğini o sutunda yer alan en  uzun yazının genişliğine göre ayarlama
            datagridview.AllowUserToAddRows = false;  //ilk sutunu gizleme
            datagridview.AllowUserToOrderColumns = true;

        }
        public void boyutAyarlama()
        {
            //girilen alternatif sayısına göre karar matrisinin boyutunu büyüten ve butonları konumlandıran kod parçası
            if (alternatifler.Count > 5)
            {
                int say = alternatifler.Count - 5;
                int sayi = 0;
                int konum, konum1, konum2 = 0;
                konum = 290 + (say * 33);
                konum1 = 313 + (say * 33);
                konum2 = 339 + (say * 33);

                sayi = 199 + (say * 33);
                if (sayi > 437)
                {
                    sayi = 437;
                }
                if (konum > 537)
                {
                    konum = 537;
                }
                if (konum1 > 560)
                {
                    konum1 = 560;
                }
                if (konum2 > 584)
                {
                    konum2 = 584;
                }
                chkPasteToSelectedCells.Top = konum;
                dataGridViewImport.Top = konum;
                label32.Top = konum1;
                btnKararMatNormalize.Top = konum2;
                btnKararMatNormalize2.Top = konum2;
                btnKararMatİleri.Top = konum2;
                dataGridViewKararMat.Height = sayi;

                konum = 260 + (say * 33);
                sayi = 199 + (say * 33);
                if (sayi > 510)
                {
                    sayi = 510;
                }
                if (konum > 583)
                {
                    konum = 583;
                }
                btnKriterAgirlikBelirleme.Top = konum;
                btnNormalizeİleri.Top = konum;
                dataGridViewNormalize.Height = sayi;


                konum = 305 + (say * 33);
                konum1 = 336 + (say * 33);
                konum2 = 279 + (say * 33);
                sayi = 199 + (say * 33);
                if (sayi > 475)
                {
                    sayi = 475;
                }
                if (konum > 581)
                {
                    konum = 581;
                }
                if (konum1 > 615)
                {
                    konum1 = 615;
                }
                if (konum2 > 555)
                {
                    konum2 = 555;
                }
                lblTetaDeğiştir.Top = konum2;
                txtTeta.Top = konum2;
                btnTetaKaydet.Top = konum2;
                label10.Top = konum;
                lblEnİyiAlternatif.Top = konum;
                btnAyrintiliCozum.Top = konum1;
                dataGridViewGenelSkor.Height = sayi;
            }
        }
        private void txtKriter_Enter(object sender, EventArgs e)
        {
            if (txtKriter.Text == "Eklemek istediğiniz kriteri giriniz")
            {
                txtKriter.Text = "";
            }
        }
        private void txtKriter_Leave(object sender, EventArgs e)
        {
            if (txtKriter.Text == "")
            {
                txtKriter.Text = "Eklemek istediğiniz kriteri giriniz";
            }
        }
        public void kriterEkle()
        {
            try
            {
                if (txtKriter.Text == "Eklemek istediğiniz kriteri giriniz")
                {
                    MessageBox.Show("Kriter girmeden ekleme yapamazsınız !", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    for (int i = 0; i < kriterler.Count; i++)
                    {
                        if (txtKriter.Text == kriterler[i].ToString())
                        {
                            MessageBox.Show("Lütfen farklı kriterler giriniz!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                    if (txtKriter.Text == "")
                    {
                        MessageBox.Show("Kriter girmeden ekleme yapamazsınız !", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    //uyarı
                    if (rbtnFayda.Checked == false && rbtnMaliyet.Checked == false)
                    {
                        MessageBox.Show("Lütfen kriter tipini seçiniz!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }


                    //listbox a ekleme
                    if (rbtnFayda.Checked == true)
                    {
                        listBoxKriter.Items.Add(txtKriter.Text + "  (" + rbtnFayda.Text + ")");
                    }
                    if (rbtnMaliyet.Checked == true)
                    {
                        listBoxKriter.Items.Add(txtKriter.Text + "  (" + rbtnMaliyet.Text + ")");
                    }

                    kriterler.Add(txtKriter.Text);
                    //eklendikten sonra butonları aktif etsin
                    btnKriterSil.Enabled = true;
                    btnKriterDuzenle.Enabled = true;

                    //fayda ve maliyet kriterlerini arrayliste ekleme

                    if (rbtnFayda.Checked == true)
                    {
                        faydaMaliyet.Add(rbtnFayda.Text);
                    }
                    if (rbtnMaliyet.Checked == true)
                    {
                        faydaMaliyet.Add(rbtnMaliyet.Text);
                    }
                    pnlKriter.Visible = true;

                    //ekledikten sonra textbox ın içini temizleyip imleci oraya fokuslasın

                    txtKriter.Clear();
                    txtKriter.Focus();

                }
            }
            catch (Exception)
            {

                MessageBox.Show("Kriter ekleme işlemi başarısız!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        private void btnKriterEkle_Click(object sender, EventArgs e)
        {

            if (btnKriterEkle.Text == "Ekle")
            {
                kriterEkle();
            }

            else if (btnKriterEkle.Text == "Güncelle")
            {
                kriterDuzenle();
            }
        }
        public void kriterDuzenle()
        {

            try
            {
                if (txtKriter.Text == "Eklemek istediğiniz kriteri giriniz")
                {
                    MessageBox.Show("Kriter girmeden ekleme yapamazsınız !", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    //for (int i = 0; i < kriterler.Count; i++)
                    //{
                    //    if (txtKriter.Text == kriterler[i].ToString())
                    //    {
                    //        MessageBox.Show("Lütfen farklı kriterler giriniz!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    //        return;
                    //    }
                    //}
                    if (txtKriter.Text == "")
                    {
                        MessageBox.Show("Kriter girmeden ekleme yapamazsınız !", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    //uyarı
                    if (rbtnFayda.Checked == false && rbtnMaliyet.Checked == false)
                    {
                        MessageBox.Show("Lütfen kriter tipini seçiniz!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }


                    //listbox a ekleme
                    if (rbtnFayda.Checked == true)
                    {
                        listBoxKriter.Items.RemoveAt(duzenleIndex);
                        listBoxKriter.Items.Insert(duzenleIndex, txtKriter.Text + "  (" + rbtnFayda.Text + ")");
                    }
                    if (rbtnMaliyet.Checked == true)
                    {
                        listBoxKriter.Items.RemoveAt(duzenleIndex);
                        listBoxKriter.Items.Insert(duzenleIndex, txtKriter.Text + "  (" + rbtnMaliyet.Text + ")");
                    }
                    kriterler.RemoveAt(duzenleIndex);
                    kriterler.Insert(duzenleIndex, txtKriter.Text);
                    //eklendikten sonra butonları aktif etsin
                    btnKriterSil.Enabled = true;
                    btnKriterDuzenle.Enabled = true;

                    //fayda ve maliyet kriterlerini arrayliste ekleme

                    if (rbtnFayda.Checked == true)
                    {
                        faydaMaliyet.RemoveAt(duzenleIndex);
                        faydaMaliyet.Insert(duzenleIndex, rbtnFayda.Text);
                    }
                    if (rbtnMaliyet.Checked == true)
                    {
                        faydaMaliyet.RemoveAt(duzenleIndex);
                        faydaMaliyet.Insert(duzenleIndex, rbtnMaliyet.Text);
                    }
                    pnlKriter.Visible = true;

                    //ekledikten sonra textbox ın içini temizleyip imleci oraya fokuslasın

                    txtKriter.Clear();
                    txtKriter.Focus();

                    btnKriterEkle.Text = "Ekle";
                    btnKriterEkle.Font = new Font("Bahnschrift Light", 9, FontStyle.Bold);

                }
            }
            catch (Exception)
            {

                MessageBox.Show("Kriter güncelleme işlemi başarısız!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public void alternatifEkle()
        {
            try
            {
                for (int i = 0; i < alternatifler.Count; i++)
                {
                    if (txtAlternatif.Text == alternatifler[i].ToString())
                    {

                        satir--;

                        MessageBox.Show("Lütfen farklı alternatifler giriniz!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
                if (txtAlternatif.Text == "Eklemek istediğiniz alternatifi giriniz")
                {

                    satir--;

                    MessageBox.Show("Alternatif girmeden ekleme yapamazsınız !", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                else if (txtAlternatif.Text == "")
                {

                    satir--;

                    MessageBox.Show("Alternatif girmeden ekleme yapamazsınız !", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    listBoxAlternatif.Items.Add(txtAlternatif.Text);
                    alternatifler.Add(txtAlternatif.Text);
                    //alternatif eklendikten sonra butonları aktif etsin
                    btnAlternatifSil.Enabled = true;
                    btnAlternatifDuzenle.Enabled = true;
                    //ekledikten sonra textbox ın içini temizleyip imleci oraya fokuslasın

                    txtAlternatif.Clear();
                    txtAlternatif.Focus();
                    pnlAlternatif.Visible = true;
                    pnlKriterAlternatif.Visible = true;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        private void btnAlternatifEkle_Click(object sender, EventArgs e)
        {
            satir++;

            if (btnAlternatifEkle.Text == "Ekle")
            {
                alternatifEkle();
            }
            else if (btnAlternatifEkle.Text == "Güncelle")
            {
                alternatifDuzenle();
            }
        }
        private void txtAlternatif_Enter(object sender, EventArgs e)
        {
            if (txtAlternatif.Text == "Eklemek istediğiniz alternatifi giriniz")
            {
                txtAlternatif.Text = "";
            }

        }
        private void txtAlternatif_Leave(object sender, EventArgs e)
        {
            if (txtAlternatif.Text == "")
            {
                txtAlternatif.Text = "Eklemek istediğiniz alternatifi giriniz";
            }
        }
        public void kararMatRenklendir()
        {
            for (int j = 1; j < kriterler.Count + 1; j++)
            {
                if (faydaMaliyet[j - 1].ToString() == rbtnFayda.Text)
                {
                    dataGridViewKararMat.Columns[j].HeaderCell.Style.BackColor = Color.LightBlue;

                }
                else if (faydaMaliyet[j - 1].ToString() == rbtnMaliyet.Text)
                {
                    dataGridViewKararMat.Columns[j].HeaderCell.Style.BackColor = Color.Plum;

                }
                else
                {
                    MessageBox.Show("boş");
                    return;
                }

            }
        }
        public void kararMatrisiOlustur()
        {
            try
            {
                //dataGridViewKararMat.Columns.Clear();
                //dataGridViewKararMat.Rows.Clear();
                tabControl1.SelectedTab = tabPageKararMatrisi;
                dataGridViewKararMat.ColumnCount = kriterler.Count + 1;
                dataGridViewKararMat.Columns[0].Name = "Virgül ile ayırınız (nokta kullanmayınız) ";
                int j = 1;
                for (int i = 0; i < kriterler.Count; i++)
                {
                    dataGridViewKararMat.Columns[j].Name = kriterler[i].ToString();
                    j++;
                }

                for (int i = alternatifler.Count - satir; i < alternatifler.Count; i++)
                {

                    dataGridViewKararMat.Rows.Add(alternatifler[i].ToString());

                }
                satir = 0;
                //İLK SUTUNDAKİ DEĞERLERİN DEĞİŞTİRİLMESİNİ ENGELLEDİM
                for (int rC = 0; rC < alternatifler.Count; rC++)
                {
                    dataGridViewKararMat.Rows[rC].Cells[0].ReadOnly = true;
                }
                kararMatRenklendir();

            }
            catch
            {

                MessageBox.Show("Karar Matrisi Oluşturulamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


        }
        public void normalizeMatCerceve()
        {
            try
            {
                dataGridViewNormalize.Columns.Clear();
                dataGridViewNormalize.Rows.Clear();
                //normalize matrisi
                dataGridViewNormalize.ColumnCount = kriterler.Count + 1; //alternatif sutunuda bulunması gerektiğinden kriterler+1 tane sutun ekledim
                dataGridViewNormalize.Columns[0].Name = " ";
                int s = 1;
                for (int i = 0; i < kriterler.Count; i++)
                {
                    dataGridViewNormalize.Columns[s].Name = kriterler[i].ToString();
                    s++;
                }
                for (int i = 0; i < alternatifler.Count; i++)
                {

                    dataGridViewNormalize.Rows.Add(alternatifler[i].ToString());
                }
                //İLK SUTUNDAKİ DEĞERLERİN DEĞİŞTİRİLMESİNİ ENGELLEDİM
                for (int rC = 0; rC < alternatifler.Count; rC++)
                {
                    dataGridViewNormalize.Rows[rC].Cells[0].ReadOnly = true;
                }


            }
            catch (Exception ex)
            {

                MessageBox.Show("Hata!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void manuelAgirlikMatrisi()
        {
            try
            {
                //eski bi rçalışma açıldığında kullanıcı ağırlık yönteminde değişiklik yapmak isteyebilir diye yöntem secme penceresini görünür yaptım
                if (tiklanma == "eski")
                {
                    pnlYontemSec.Visible = true;
                }
                else if (tiklanma != "eski")
                {
                    pnlYontemSec.Visible = false;
                }

                panel12.Visible = true;
                dataGridViewAgirlik.Columns.Clear();
                dataGridViewAgirlik.Rows.Clear();
                int k = 1;
                dataGridViewAgirlik.ColumnCount = kriterler.Count + 1;
                dataGridViewAgirlik.Columns[0].Name = " ";

                for (int j = 0; j < kriterler.Count; j++)
                {
                    dataGridViewAgirlik.Columns[k].Name = kriterler[j].ToString();
                    k++;
                }
                for (int j = 0; j < 1; j++)
                {
                    dataGridViewAgirlik.Rows.Add("Ağırlıklar (virgül ile ayırınız)");
                }

                dataGridViewAgirlik.Rows[0].Cells[0].ReadOnly = true;
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        private void btnManuel_Click(object sender, EventArgs e)
        {
            yontem = "manuel";
            lblKarMat.Visible = false;
            agirlikTemizle();
            manuelAgirlikMatrisi();

        }
        private void btnAgirlikKaydet_Click(object sender, EventArgs e)
        {
            sonuçlarToolStripMenuItem.Visible = true;
            try
            {
                pnlYontemSec.Visible = true;
                agirliklar.Clear();
                //ağırlıklar listesine ağırlık değerlerini ekledim
                for (int j = 1; j < kriterler.Count + 1; j++)
                {
                    agirliklar.Add(dataGridViewAgirlik.Rows[0].Cells[j].Value.ToString());
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //AĞIRLIK TOPLAMALARININ 1 EŞİT OLUP OLMADIĞININ KONTROLÜ
            agirlikToplam = 0;
            foreach (var item in agirliklar)
            {
                agirlikToplam += Convert.ToDouble(item);
            }
            if (agirlikToplam > 1 || agirlikToplam < 0.99)
            {
                MessageBox.Show("Toplam ağırlık: " + agirlikToplam + " Lütfen girilen değerleri kontrol edip tekrar deneyiniz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            goreliAgirliklarim();
            maxGenelBakinlik();
            genelBaskinlikSkoruMatrisi();
            tabControl1.SelectedTab = tabPageGenelBaskinlik;
        }
        //en yülsek ağırlığa sahip değeri referans Kriteri olarak belirledim
        public void referansKriteri()
        {
            try
            {
                wr = Convert.ToDouble(agirliklar[0]);

                foreach (var sayi in agirliklar)
                {
                    if (Convert.ToDouble(sayi) > wr)
                    {
                        wr = Convert.ToDouble(sayi);
                    }

                }
            }
            catch (Exception)
            {

                MessageBox.Show("Referans kriteri hesaplanamadı. Lütfen ağırlık değerlerini kontrol ediniz!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void goreliAgirliklarim()
        {
            try
            {
                referansKriteri();
                goreliAgirliklar.Clear(); // teta için ekledim
                //goreli ağırlıklar
                for (int j = 0; j < kriterler.Count; j++)
                {
                    wjr = Convert.ToDouble(agirliklar[j]) / wr;
                    goreliAgirliklar.Add(wjr);
                }
                //göreli ağırlık toplamı

                foreach (var wjr in goreliAgirliklar)
                {
                    gAgirlikTop += Convert.ToDouble(wjr);
                }
                goreliAgirliklar.Add(gAgirlikTop);
                gAgirlikTop = 0; // teta için ekledim
            }
            catch (Exception)
            {

                MessageBox.Show("Göreli ağırlıklar hesaplanamadı. Lütfen girilen değerleri kontrol ediniz!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        private void btnKriterDuzenle_Click(object sender, EventArgs e)
        {
            duzenleIndex = listBoxKriter.SelectedIndex;
            txtKriter.Text = kriterler[duzenleIndex].ToString();
            btnKriterEkle.Text = "Güncelle";
            btnKriterEkle.Font = new Font("Bahnschrift Light", 8, FontStyle.Bold);
        }
        public void kriterSil()
        {
            try
            {
                int secili = listBoxKriter.SelectedIndex;
                faydaMaliyet.RemoveAt(secili);
                kriterler.RemoveAt(secili);
                listBoxKriter.Items.RemoveAt(secili);

                if (listBoxKriter.Items.Count == 0)
                {
                    txtKriter.Focus();
                    btnKriterSil.Enabled = false;

                }
            }
            catch (Exception)
            {
                MessageBox.Show("Lütfen silmek istediğiniz kriteri seçiniz!", "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        private void btnKriterSil_Click(object sender, EventArgs e)
        {
            if (dataGridViewKararMat.Columns.Count != 0)
            {
                dataGridViewKararMat.Columns.RemoveAt(listBoxKriter.SelectedIndex + 1);
            }
            kriterSil(); //kriteri kriterler arraylistinden silen method 
        }
        private void btnAlternatifSil_Click(object sender, EventArgs e)
        {
            if (dataGridViewKararMat.Rows.Count != 0)
            {
                dataGridViewKararMat.Rows.RemoveAt(listBoxAlternatif.SelectedIndex);
            }
            alternatifSil(); //alternatifi alternatifler listesinden silen method          
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            bilgiMesaji("Excel dosyası indirme", "Örnek Excel şablonunu indirmek için tıklayınız.", btnOrnekExcelDosya);
            bilgiMesaji("Karar matrisi oluşturma", "Karar matrisini oluşturmak için tıklayınız.", btnSimdiOlustur);
            bilgiMesaji("Karar matrisi oluşturma", "Karar matrisini excel'den yüklemek için tıklayınız.", btnExcelYukle);
            if (btnKriterEkle.Text == "Güncelle")
            {
                bilgiMesaji("Kriter güncelleme", "Seçilen kriteri güncellemek için tıklayınız.", btnKriterEkle);

            }

            else if (btnKriterEkle.Text == "Ekle")
            {
                bilgiMesaji("Karar matrisi oluşturma", "Kriter eklemek için tıklayınız.", btnKriterEkle);

            }
            if (btnAlternatifEkle.Text == "Güncelle")
            {
                bilgiMesaji("Alternatif güncelleme", "Seçilen alternatifi güncellemek için tıklayınız.", btnAlternatifEkle);

            }
            else if (btnAlternatifEkle.Text == "Ekle")
            {
                bilgiMesaji("Karar matrisi oluşturma", "Alternatif eklemek için tıklayınız.", btnAlternatifEkle);

            }

            bilgiMesaji("Kriter yönü", "Fayda Kriteri.", rbtnFayda);
            bilgiMesaji("Kriter yönü", "Maliyet Kriteri.", rbtnMaliyet);
            bilgiMesaji("Kriter düzenleme", "Seçilen kriteri düzenlemek için tıklayınız.", btnKriterDuzenle);
            bilgiMesaji("Kriter silme", "Seçilen kriteri silmek için tıklayınız.", btnKriterSil);
            bilgiMesaji("Alternatif düzenleme", "Seçilen alternatifi düzenlemek için tıklayınız.", btnAlternatifDuzenle);
            bilgiMesaji("Alternatif silme", "Seçilen alternatifi silmek için tıklayınız.", btnAlternatifSil);
            bilgiMesaji("Karar matrisi oluşturma", "Karar matrisini oluşturmak için tıklayınız.", btnKararMatOL);
            bilgiMesaji("Normalizasyon", "Karar matrisini min max normalizasyon yöntemini kullanarak normalize etmek için tıklayınız.", btnKararMatNormalize);
            bilgiMesaji("Normalizasyon", "Karar matrisini normalize etmek için tıklayınız.", btnKararMatNormalize2);
            bilgiMesaji("Excel aktarma", "Karar matrisini excel'e aktarmak için tıklayınız.", btnKararMatEAktar);
            bilgiMesaji("Excel'de açma", "Karar matrisini excelde açmak  için tıklayınız.", btnKararMatExceldeAc);
            bilgiMesaji("Kriter ağırlığı belirleme", "Kriter ağırlıklarını belirlemek için tıklayınız.", btnKriterAgirlikBelirleme);
            bilgiMesaji("Kriter ağırlığı belirleme", "Kriter ağırlıklarını el ile girmek için tıklayınız.", btnManuel);
            bilgiMesaji("Kriter ağırlığı belirleme", "Kriter ağırlıklarını AHP ile hesaplatmak için tıklayınız.", btnAhp);
            bilgiMesaji("Ağırlık değerlerini kaydet", "Girilen ağırlık değerlerini kaydetmek için tıklayınız.", btnAgirlikKaydet);
            bilgiMesaji("Ağırlık hesaplatma", "AHP ile ağırlık değerlerini hesaplatmak için tıklayınız.", btnAhpHesapla);
            bilgiMesaji("Ağırlık değerlerini düzenleme", "İkili karşılaştırma matrisi oluşturmaya dönmek için tıklayınız.", btnAgirlikDuzenle);
            bilgiMesaji("Ağırlık değerlerini kaydetme", "Genel baskınlık skorlarını görüntülemek için tıklayınız.", btnAhpAyrintiAgirlikKaydet);
            bilgiMesaji("Ağırlık değerlerini kaydetme", "Genel baskınlık skorlarını görüntülemek için tıklayınız.", btnKriterAgirlikKaydet);
            bilgiMesaji("AHP ayrıntılı çözümler", "AHP ayrıntılı çözümleri görüntülemek için tıklayınız.", btnAhpAyrintiCozGoster);
            bilgiMesaji("Excel'e aktarma", "AHP ayrıntılı çözümleri excel'e aktarmak için tıklayınız.", btnAhpKriterAğrBulEAktar);
            bilgiMesaji("Tutarlılık oranı", "Tutarlılık oranı %10'un altında olmalıdır.", label18);
            bilgiMesaji("Kriter ağırlıkları", "", label20);
            bilgiMesaji("θ değerini değiştir", "θ değerini değiştirmek için tıklayınız.", lblTetaDeğiştir);
            bilgiMesaji("θ değerini değiştir", "Girilen θ değerine göre hesaplama yaptırmak için tıklayınız.", btnTetaKaydet);
            bilgiMesaji("Ağırlık değerlerini kaydetme", "Genel baskınlık skorlarını görüntülemek için tıklayınız.", btnKriterAgirlikKaydet);
            bilgiMesaji("Ayrıntılı çözümler", "Ayrıntılı çözümleri görüntülemek için tıklayınız.", btnAyrintiliCozum);
            bilgiMesaji("Excel'e aktarma", "Genel baskınlık skorlarını excel'e aktarmak için tıklayınız.", btnGBS_eAktar);
            bilgiMesaji("Bakınlık skoru arama", "Baskınlık skorunu görüntülemek için tıklayınız.", btnBaskinlikSkoru);
            bilgiMesaji("Excel'de aç", "Baskınlık skorunu excel'de açmak için tıklayınız.", btnBSkorEAc);
            bilgiMesaji("Excel'e aktarma", "Baskınlık skorunu excel'e aktarmak için tıklayınız.", btnBSkorEAktar);
            bilgiMesaji("Excel'e aktarma", "Kısmi baskınlık skorlarını excel'e aktarmak için tıklayınız.", btnTumKismiBskorExcelAktar);
            bilgiMesaji("Kısmi baskınlık skoru görüntüleme", "Tüm kısmi baskınlık skorlarını görüntülemek için tıklayınız.", btnTumBSkor);
            bilgiMesaji("Kaydet", "Çalışmanızı kaydetmek için tıklayınız.", btnCalismayiKaydet);
            bilgiMesaji("Excel'e aktarma", "Sonuçları excel'e aktarmak için tıklayınız.", btnSonucExcelAktar);




            //  TABCONTROL BAŞLIK GİZLEME

            //Rectangle rect = new Rectangle(
            //tabPageBaslangic.Left,
            //tabPageBaslangic.Top,
            //tabPageBaslangic.Width,
            //tabPageBaslangic.Height);
            //tabControl1.Region = new Region(rect);
        }
        private void btnAhpHesapla_Click(object sender, EventArgs e)
        {
            try
            {
                btnIkiliKMatEAc.Visible = true;
                btnIkiliKMatEAktar.Visible = true;
                btnKriterAgirlikKaydet.Visible = true;
                btnAhpAyrintiCozGoster.Visible = true;
                dataGridViewKriterAgirliklari.Rows.Clear();
                panel82.Visible = true;
                panel81.Visible = true;
                panel83.Visible = true;
                dgvKriterAgirlikDoldur();
                panel80.Visible = true;
                label24.Visible = true;
                lblTutOrani.Visible = true;


                if (Convert.ToDouble(lblTutOrani.Text) >= 0.10)
                {
                    lblUyari.Visible = true;
                    panel52.Visible = true;
                    btnKriterAgirlikKaydet.Text = "Bu şekilde devam et";
                }


            }
            catch (Exception)
            {

                MessageBox.Show("AHP Hesaplanamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void dgvKriterAgirlikDoldur()
        {
            try
            {
                dataGridViewKriterAgirliklari.ColumnCount = dataGridViewWVektörü.Columns.Count;
                int k = 1;
                for (int j = 0; j < kriterler.Count; j++)
                {
                    dataGridViewKriterAgirliklari.Columns[k].Name = kriterler[j].ToString();
                    k++;
                }
                dataGridViewKriterAgirliklari.Rows.Add("Ağırlıklar");
                for (int i = 0; i < dataGridViewWVektörü.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridViewWVektörü.Columns.Count; j++)
                    {
                        dataGridViewKriterAgirliklari.Rows[i].Cells[j].Value = dataGridViewWVektörü.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Kriter ağırlıkları getirilemedi!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void alternatifSil()
        {
            try
            {
                int secili = listBoxAlternatif.SelectedIndex;
                alternatifler.RemoveAt(secili);
                listBoxAlternatif.Items.RemoveAt(secili);
                txtAlternatif.Focus();
                if (listBoxAlternatif.Items.Count == 0)
                {
                    btnAlternatifSil.Enabled = false;
                    txtAlternatif.Focus();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Lütfen silmek istediğiniz alternatifi seçiniz!", "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        private void dataGridViewKarsilastirmaMat_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int y = dataGridViewKarsilastirmaMat.CurrentCell.ColumnIndex;  //seçili sutunun indexini tutar
                int x = dataGridViewKarsilastirmaMat.CurrentCell.RowIndex + 1; // seçili satır indexinin 1 fazlasını tutar
                if (y != 0 && x != y)
                {
                    dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value = 1 / Convert.ToDouble(dataGridViewKarsilastirmaMat.Rows[x - 1].Cells[y].Value);
                    dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value = Math.Round(Convert.ToDouble(dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value), 2);
                }
                if (y != 0 && x == y)
                {
                    dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value = Convert.ToDouble(dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value) / Convert.ToDouble(dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value);
                    dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value = Math.Round(Convert.ToDouble(dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value), 2);
                }
                if (dataGridViewKarsilastirmaMat.Rows[x - 1].Cells[y].Value != null)
                {
                    if (y != 0 && Convert.ToDouble(dataGridViewKarsilastirmaMat.Rows[x - 1].Cells[y].Value) > 9)
                    {
                        MessageBox.Show("Lütfen önem değerlerini 1-9 arasında puanlayın!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        dataGridViewKarsilastirmaMat.Rows[x - 1].Cells[y].Value = 1;
                        dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value = 1;
                    }
                    if (y != 0 && Convert.ToDouble(dataGridViewKarsilastirmaMat.Rows[x - 1].Cells[y].Value) <= 0)
                    {
                        MessageBox.Show("Lütfen önem değerlerini 1-9 arasında puanlayın!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        dataGridViewKarsilastirmaMat.Rows[x - 1].Cells[y].Value = 1;
                        dataGridViewKarsilastirmaMat.Rows[y - 1].Cells[x].Value = 1;
                    }
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Lütfen yalnızca sayısal değerler girin!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        public void eskiCalismaEkleSilGüncelle()
        {
            //karar matrisi
            if (alternatifler.Count > 1 && kriterler.Count > 1)
            {
                boyutAyarlama();
                dataGridViewKararMat.Columns.Clear();
                dataGridViewKararMat.Rows.Clear();
                kararMatrisiOlustur();
            }
            else
            {
                MessageBox.Show("Karar matrisini oluşturmak için yeterli alternatif veya kriteriniz yok!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //NORMALİZE
            if (secilenNormalizeYontemi == "1")
            {

                normalizeEdilmisKararMatrisiToolStripMenuItem.Visible = true;
                normalizeMatCerceve();
                //boş hücre kontrolü
                for (int i = 0; i < alternatifler.Count; i++) //satır
                {
                    for (int j = 1; j < kriterler.Count + 1; j++) //sutun
                    {
                        if (dataGridViewKararMat.Rows[i].Cells[j].Value == null)
                        {
                            MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                    }
                }
                minMaxNormalization();

                if (tiklanma == "eski")
                {
                    pnlYontemSec.Visible = true;
                }

                tabControl1.SelectedTab = tabPageNormalize;

            }
            else if (secilenNormalizeYontemi == "2")
            {
                normalizeEdilmisKararMatrisiToolStripMenuItem.Visible = true;
                normalizeMatCerceve();
                for (int i = 0; i < alternatifler.Count; i++) //satır
                {
                    for (int j = 1; j < kriterler.Count + 1; j++) //sutun
                    {
                        if (dataGridViewKararMat.Rows[i].Cells[j].Value == null)
                        {
                            MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                normalizeYeni();

                if (tiklanma == "eski")
                {
                    pnlYontemSec.Visible = true;
                }
            }



        }
        private void btnKararMatOL_Click(object sender, EventArgs e)
        {
            if (tiklanma == "olustur" || tiklanma == "excel")
            {
                if (alternatifler.Count > 1 && kriterler.Count > 1)
                {
                    boyutAyarlama();
                    kararMatrisiOlustur();
                }
                else
                {
                    MessageBox.Show("Karar matrisini oluşturmak için yeterli alternatif veya kriteriniz yok!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }


            else if (tiklanma == "eski")
            {
                MessageBox.Show("Eski çalışma için kriter ve alternatiflerde henüz değişiklik yapılamammaktadır!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }



        }
        private void btnBaskinlikSkoru_Click(object sender, EventArgs e)
        {
            dataGridViewBaskinlikSkoru.Rows.Clear();
            dataGridViewBaskinlikSkoru.Columns.Clear();
            panel10.Visible = true;
            btnBSkorEAktar.Visible = true;
            pjDegerleri();
        }
        private void btnKriterAgirlikBelirleme_Click(object sender, EventArgs e)
        {
            if (tiklanma == "eski")
            {
                pnlYontemSec.Visible = true;
            }
            agirlikYontemSecToolStripMenuItem.Visible = true;
            kriterAgirligiBelirlemeToolStripMenuItem.Visible = true;
            tabControl1.SelectedTab = tabPageAgirlikBelirleme;
        }
        public void ahpAgirlikMatrisi()
        {
            dataGridViewKarsilastirmaMat.Columns.Clear();
            dataGridViewKarsilastirmaMat.Rows.Clear();
            pnlAhpTasarim.Controls.Clear();
            int k = 1;
            dataGridViewKarsilastirmaMat.ColumnCount = kriterler.Count + 1;
            dataGridViewKarsilastirmaMat.Columns[0].Name = " ";

            for (int i = 0; i < kriterler.Count; i++)
            {
                k = 1;
                for (int j = 0; j < kriterler.Count; j++)
                {
                    dataGridViewKarsilastirmaMat.Columns[k].Name = kriterler[j].ToString();
                    k++;
                }
            }
            for (int i = 0; i < kriterler.Count; i++)
            {
                k = 1;
                for (int j = 0; j < kriterler.Count; j++)
                {
                    dataGridViewKarsilastirmaMat.Columns[k].Name = kriterler[j].ToString();
                    k++;
                }
            }
            for (int j = 0; j < kriterler.Count; j++)
            {
                dataGridViewKarsilastirmaMat.Rows.Add(kriterler[j].ToString());
            }
            for (int cR = 0; cR < kriterler.Count; cR++)
            {
                dataGridViewKarsilastirmaMat.Rows[cR].Cells[0].ReadOnly = true;
            }
            for (int i = 0; i < kriterler.Count; i++)
            {
                dataGridViewKarsilastirmaMat.Rows[i].Cells[i + 1].Value = 1;
            }

        }
        private void btnAhp_Click(object sender, EventArgs e)
        {
           
            try
            {
                yontem = "ahp";
                AHPToolStripMenuItem.Visible = true;
                karsilastirmaMatrisiOlusturmaToolStripMenuItem.Visible = true;
                agirlikTemizle();
                ahpAgirlikMatrisi();
                lblKarMat.Visible = true;
                ahpTasarim();
                tabControl1.SelectedTab = tabPageKarsilastirmaMat;
            }
            catch (Exception ex)
            {
                MessageBox.Show("AHP Hesaplanamadı! " + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void bosHucreKontrolu() // karar matrisi için
        {
            for (int i = 0; i < alternatifler.Count; i++) //satır
            {
                for (int j = 1; j < kriterler.Count + 1; j++) //sutun
                {
                    if (dataGridViewKararMat.Rows[i].Cells[j].Value == null)
                    {
                        MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
            }
        }
        //min max normalizasyon
        public void minMaxNormalization() //karar matrisini fayda ve maliyet kriteri olmasına göre ayrı formüllerle normalize edip normalizasyon matrisini oluşturan metod
        {
            try
            {
                maxMin(); //her sutundaki max ve min değerleri bulup ilgili listelere atayan metod
                //normalizasyon
                double max1, min1;
                for (int j = 1; j < kriterler.Count + 1; j++)
                {
                    if (faydaMaliyet[j - 1].ToString() == "Fayda")
                    {
                        max1 = Convert.ToDouble(maxList[j - 1]);
                        min1 = Convert.ToDouble(minList[j - 1]);

                        for (int i = 0; i < alternatifler.Count; i++)
                        {
                            if (max1 - min1 == 0)
                            {
                                dataGridViewNormalize.Rows[i].Cells[j].Value = 0;
                            }
                            else
                            {
                                dataGridViewNormalize.Rows[i].Cells[j].Value = (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value) - min1) / (max1 - min1);

                            }
                        }
                    }
                    else /*if (faydaMaliyet[j - 1].ToString() == "Maliyet")*/
                    {
                        max1 = Convert.ToDouble(maxList[j - 1]);
                        min1 = Convert.ToDouble(minList[j - 1]);
                        for (int i = 0; i < alternatifler.Count; i++)
                        {
                            if (max1 - min1 == 0)
                            {
                                dataGridViewNormalize.Rows[i].Cells[j].Value = 0;
                            }
                            else
                            {
                                dataGridViewNormalize.Rows[i].Cells[j].Value = (max1 - (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value))) / (max1 - min1);
                            }

                        }

                    }
                }
                tabControl1.SelectedTab = tabPageNormalize;// butona tıklanıldığında tabPageNormalize ye gönderen kod
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "  Matris Normalize Edilemedi! Lütfen metinsel değerler girmeyiniz.", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public void kararMatSutunTopla()
        {
            //sutun toplamlarını hesaplayıp kararMatSutunToplam adındaki arrayliste ekledim

            for (int j = 1; j < kriterler.Count + 1; j++)
            {
                if (faydaMaliyet[j - 1].ToString() == "Fayda")
                {
                    double sutunToplam = 0;
                    for (int i = 0; i < alternatifler.Count; i++)
                    {
                        sutunToplam += Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value);
                    }
                    kararMatSutunToplam.Add(sutunToplam);
                }
                else
                {
                    double sutunBireBolTopla = 0;
                    for (int i = 0; i < alternatifler.Count; i++)
                    {
                        sutunBireBolTopla += Convert.ToDouble(1 / (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value)));
                    }
                    kararMatSutunToplam.Add(sutunBireBolTopla);
                }
            }

        }
        public void normalizeYeni() //karar matrisini fayda ve maliyet kriteri olmasına göre ayrı formüllerle normalize edip normalizasyon matrisini oluşturan metod
        {
            try
            {
                for (int j = 1; j < kriterler.Count + 1; j++)
                {
                    if (faydaMaliyet[j - 1].ToString() == "Fayda")
                    {
                        double sutunToplam = 0;
                        for (int i = 0; i < alternatifler.Count; i++)
                        {
                            sutunToplam += Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value);
                        }
                        kararMatSutunToplam.Add(sutunToplam);
                    }
                    else
                    {
                        double sutunBireBolTopla = 0;
                        for (int i = 0; i < alternatifler.Count; i++)
                        {
                            sutunBireBolTopla += Convert.ToDouble(1 / (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value)));
                        }
                        kararMatSutunToplam.Add(sutunBireBolTopla);
                    }
                }

                //normalizasyon

                for (int j = 1; j < kriterler.Count + 1; j++)
                {
                    for (int i = 0; i < alternatifler.Count; i++)
                    {
                        double sayi = Convert.ToDouble(kararMatSutunToplam[j - 1]);
                        if (faydaMaliyet[j - 1].ToString() == "Fayda")
                        {
                            if (Convert.ToDouble(kararMatSutunToplam[j - 1]) == 0)
                            {
                                dataGridViewNormalize.Rows[i].Cells[j].Value = 0;
                            }
                            else
                            {
                                dataGridViewNormalize.Rows[i].Cells[j].Value = (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value)) / sayi;
                            }
                        }
                        else /*if (faydaMaliyet[j - 1].ToString() == "Maliyet")*/
                        {
                            if (Convert.ToDouble(kararMatSutunToplam[j - 1]) == 0)
                            {
                                dataGridViewNormalize.Rows[i].Cells[j].Value = 0;
                            }
                            else
                            {
                                dataGridViewNormalize.Rows[i].Cells[j].Value = (1 / (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value))) / sayi;
                            }
                        }
                    }
                }
                tabControl1.SelectedTab = tabPageNormalize; // butona tıklanıldığında tabPageNormalize ye gönderen metod
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "  Matris Normalize Edilemedi! Lütfen metinsel değerler girmeyiniz.", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void btnKararMatNormalize_Click(object sender, EventArgs e)
        {


            secilenNormalizeYontemi = "1";

            normalizeEdilmisKararMatrisiToolStripMenuItem.Visible = true;
            normalizeMatCerceve();
            //boş hücre kontrolü
            for (int i = 0; i < alternatifler.Count; i++) //satır
            {
                for (int j = 1; j < kriterler.Count + 1; j++) //sutun
                {
                    if (dataGridViewKararMat.Rows[i].Cells[j].Value == null)
                    {
                        MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                }
            }
            minMaxNormalization();

            if (tiklanma == "eski")
            {
                pnlYontemSec.Visible = true;
            }
        }
        public void maxMin() //KARAR MATRİSİNDEKİ HER SUTUNDAKİ MAX VE MİN DEĞERLERİ BULUP ARRAYLİSTLERE ATAN METOD
        {
            //karar matrisindeki bir sutundaki max ve min değerleri bulan döngüler
            for (int j = 1; j < kriterler.Count + 1; j++)
            {
                max = Convert.ToDouble(dataGridViewKararMat.Rows[0].Cells[j].Value);
                min = Convert.ToDouble(dataGridViewKararMat.Rows[0].Cells[j].Value);

                for (int i = 0; i < alternatifler.Count; i++)
                {
                    if (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value) > max)
                    {
                        max = Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value);
                    }
                }
                maxList.Add(max);

                for (int i = 0; i < alternatifler.Count; i++)
                {
                    if (Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value) < min)
                    {
                        min = Convert.ToDouble(dataGridViewKararMat.Rows[i].Cells[j].Value);
                    }
                }
                minList.Add(min);
            }
        }
        public void pjMatrisi()
        {

            dataGridViewBaskinlikSkoru.Rows.Clear();
            dataGridViewBaskinlikSkoru.Columns.Clear();
            alternatifBskor = alternatifler[0].ToString();
            dataGridViewBaskinlikSkoru.ColumnCount = kriterler.Count + 2;
            dataGridViewBaskinlikSkoru.Columns[0].Name = " ";
            dataGridViewBaskinlikSkoru.Columns[kriterler.Count + 1].Name = "Toplam";
            int j = 1;
            for (int i = 0; i < kriterler.Count; i++)
            {
                dataGridViewBaskinlikSkoru.Columns[j].Name = "P" + (i + 1).ToString() + "(Ai,Ai')";
                j++;
            }

            for (int i = 0; i < alternatifler.Count; i++)
            {
                alternatifBskor = alternatifler[i].ToString();
                girilenAlternatif = txtGirilenAlternatif.Text;
                try
                {
                    if (girilenAlternatif == alternatifBskor)
                    {
                        dataGridViewBaskinlikSkoru.Rows.Clear();

                        for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                        {
                            if (iUssu == i)
                            {
                                continue;
                            }

                            dataGridViewBaskinlikSkoru.Rows.Add(("(" + alternatifler[i].ToString()) + "," + alternatifler[iUssu].ToString() + ")");

                        }
                    }
                    else
                    {
                        continue;
                    }

                }
                catch (Exception)
                {

                    MessageBox.Show("Böyle bir alternatif bulunmamaktadır.", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

            }

            //ilk sutundaki değerlerin değiştirilmesini önleme
            for (int rC = 0; rC < alternatifler.Count; rC++)
            {
                dataGridViewBaskinlikSkoru.Rows[rC].Cells[0].ReadOnly = true;
            }


        }
        public void pjDegerleri() // Ai. alternatifin Ai' alternatifine baskınlık skoruna j. kriterin katkısını hesaplayan metod 
        {
            try
            {
                string bulunduMu = "";
                //pjMatrisi();
                foreach (var item in alternatifler)
                {
                    if (txtGirilenAlternatif.Text == item.ToString())
                    {
                        bulunduMu = "bulundu";
                    }
                }
                if (bulunduMu == "bulundu")
                {
                    for (int i = 0; i < alternatifler.Count; i++)
                    {
                        string a = alternatifler[i].ToString();
                        if (txtGirilenAlternatif.Text == a)
                        {
                            dataGridViewBaskinlikSkoru.ColumnCount = kriterler.Count + 2;
                            dataGridViewBaskinlikSkoru.Columns[0].Name = " ";
                            dataGridViewBaskinlikSkoru.Columns[kriterler.Count + 1].Name = "Toplam";
                            int j = 1;
                            for (int ii = 0; ii < kriterler.Count; ii++)
                            {
                                dataGridViewBaskinlikSkoru.Columns[j].Name = "P" + (ii + 1).ToString() + "(Ai,Ai')";
                                j++;
                            }

                            for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                            {
                                if (iUssu == i)
                                {
                                    continue;
                                }

                                dataGridViewBaskinlikSkoru.Rows.Add(("(" + alternatifler[i].ToString()) + "," + alternatifler[iUssu].ToString() + ")");

                            }


                            //label7.Text = alternatifBskor + " ALTERNATİFİNİN " + alternatifBskor + "' ALTERNATİFLERİNE BASKINLIK SKORU:";
                            for (int k = 1; k < kriterler.Count + 1; k++)
                            {
                                for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                                {
                                    if (iUssu == i)
                                    {
                                        continue;
                                    }

                                    rijFark = Convert.ToDouble(dataGridViewNormalize.Rows[i].Cells[k].Value) - Convert.ToDouble(dataGridViewNormalize.Rows[iUssu].Cells[k].Value);
                                    wjr1 = Convert.ToDouble(goreliAgirliklar[k - 1]);
                                    int indis = kriterler.Count;
                                    wjrToplam = Convert.ToDouble(goreliAgirliklar[indis]);

                                    if (rijFark > 0)
                                    {
                                        if (wjrToplam == 0)
                                        {
                                            sonuc = (wjr1 * rijFark);
                                            sonuc = Math.Sqrt(sonuc);

                                            if (i > iUssu)
                                            {
                                                dataGridViewBaskinlikSkoru.Rows[iUssu].Cells[k].Value = sonuc;
                                            }
                                            else if (i < iUssu)
                                            {
                                                dataGridViewBaskinlikSkoru.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                            }
                                        }

                                        else
                                        {
                                            sonuc = (wjr1 * rijFark) / wjrToplam;
                                            sonuc = Math.Sqrt(sonuc);
                                            if (i > iUssu)
                                            {
                                                dataGridViewBaskinlikSkoru.Rows[iUssu].Cells[k].Value = sonuc;
                                            }
                                            else if (i < iUssu)
                                            {
                                                dataGridViewBaskinlikSkoru.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                            }
                                        }

                                    }

                                    else if (rijFark == 0)
                                    {

                                        sonuc = 0;
                                        if (i > iUssu)
                                        {
                                            dataGridViewBaskinlikSkoru.Rows[iUssu].Cells[k].Value = sonuc;
                                        }
                                        else if (i < iUssu)
                                        {
                                            dataGridViewBaskinlikSkoru.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                        }

                                    }

                                    else //rijFark <0 durumu
                                    {

                                        if (wjr1 == 0)
                                        {
                                            if (θ == 0)
                                            {
                                                sonuc = 0;

                                                if (i > iUssu)
                                                {
                                                    dataGridViewBaskinlikSkoru.Rows[iUssu].Cells[k].Value = sonuc;
                                                }
                                                else if (i < iUssu)
                                                {
                                                    dataGridViewBaskinlikSkoru.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                                }
                                            }
                                            else
                                            {
                                                sonuc = (wjrToplam * rijFark) * -1;
                                                sonuc = (Math.Sqrt(sonuc)) * -Convert.ToDouble(1 / θ);


                                                if (i > iUssu)
                                                {
                                                    dataGridViewBaskinlikSkoru.Rows[iUssu].Cells[k].Value = sonuc;
                                                }
                                                else if (i < iUssu)
                                                {
                                                    dataGridViewBaskinlikSkoru.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                                }
                                            }
                                        }

                                        else
                                        {
                                            if (θ == 0)
                                            {
                                                sonuc = 0;

                                                if (i > iUssu)
                                                {
                                                    dataGridViewBaskinlikSkoru.Rows[iUssu].Cells[k].Value = sonuc;
                                                }
                                                else if (i < iUssu)
                                                {
                                                    dataGridViewBaskinlikSkoru.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                                }
                                            }
                                            else
                                            {
                                                sonuc = ((wjrToplam * rijFark) / wjr1) * -1;
                                                sonuc = (Math.Sqrt(sonuc)) * -Convert.ToDouble(1 / θ);

                                                if (i > iUssu)
                                                {
                                                    dataGridViewBaskinlikSkoru.Rows[iUssu].Cells[k].Value = sonuc;
                                                }
                                                else if (i < iUssu)
                                                {
                                                    dataGridViewBaskinlikSkoru.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                                }
                                            }
                                        }

                                    }
                                }

                            }

                            for (int s = 0; s < alternatifler.Count - 1; s++)
                            {
                                for (int c = 1; c < kriterler.Count + 1; c++)
                                {
                                    satirSkorToplam += Convert.ToDouble(dataGridViewBaskinlikSkoru.Rows[s].Cells[c].Value);
                                }
                                dataGridViewBaskinlikSkoru.Rows[s].Cells[kriterler.Count + 1].Value = satirSkorToplam;
                                baskinlikSkoru += satirSkorToplam;
                                satirSkorToplam = 0;
                            }

                            dataGridViewBaskinlikSkoru.Rows.Add("S(Ai,Ai')");
                            dataGridViewBaskinlikSkoru.Rows[alternatifler.Count - 1].Cells[kriterler.Count + 1].Value = baskinlikSkoru;

                            baskinlikSkoru = 0;

                        }

                        else
                        {
                            continue;
                        }
                    }
                }
                else if (bulunduMu != "bulundu")
                {
                    dataGridViewBaskinlikSkoru.Rows.Clear();
                    dataGridViewBaskinlikSkoru.Columns.Clear();
                    btnBSkorEAktar.Visible = false;
                    panel10.Visible = false;
                    //panel11.Visible = false;
                    MessageBox.Show("Böyle bir alternatif bulunmamaktadır! Lütfen girdiğiniz değerleri kontrol ediniz.", "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

            }
            catch (Exception ex)
            {
                dataGridViewBaskinlikSkoru.Rows.Clear();
                dataGridViewBaskinlikSkoru.Columns.Clear();
                btnBSkorEAktar.Visible = false;
                panel10.Visible = false;
                //panel11.Visible = false;
                MessageBox.Show("Böyle bir alternatif bulunmamaktadır! Lütfen girdiğiniz değerleri kontrol ediniz." + ex.Message, "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        //Tüm baskınlık skorlarını hesaplayıp baskinlikSkorlari listesine ekleyen metod
        public void tümBaskinlikSkorlari() // Ai. alternatifin Ai' alternatifine baskınlık skoruna j. kriterin katkısını hesaplayan metod 
        {
            try
            {
                tBaskinlikSkorlari.Clear(); //teta için ekledim
                tbaskinlikSkoru = 0;//teta için ekledim 
                for (int i = 0; i < alternatifler.Count; i++)
                {
                    for (int j = 1; j < kriterler.Count + 1; j++)
                    {
                        for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                        {
                            if (iUssu == i)
                            {
                                continue;
                            }

                            rijFark = Convert.ToDouble(dataGridViewNormalize.Rows[i].Cells[j].Value) - Convert.ToDouble(dataGridViewNormalize.Rows[iUssu].Cells[j].Value);
                            wjr1 = Convert.ToDouble(goreliAgirliklar[j - 1]);
                            int indis = kriterler.Count;
                            wjrToplam = Convert.ToDouble(goreliAgirliklar[indis]);

                            if (rijFark > 0)
                            {
                                if (wjrToplam == 0)
                                {
                                    sonuc1 = (wjr1 * rijFark);
                                    sonuc1 = Math.Sqrt(sonuc1);
                                    tbaskinlikSkoru += sonuc1;
                                }

                                else
                                {
                                    sonuc1 = (wjr1 * rijFark) / wjrToplam;
                                    sonuc1 = Math.Sqrt(sonuc1);
                                    tbaskinlikSkoru += sonuc1;
                                }
                            }

                            else if (rijFark == 0)
                            {
                                sonuc1 = 0;
                                tbaskinlikSkoru += sonuc1;
                            }

                            else //rijFark <0 durumu
                            {
                                if (θ == 0)
                                {
                                    tbaskinlikSkoru += 0;
                                }
                                else
                                {
                                    if (wjr1 == 0)
                                    {

                                        sonuc1 = (wjrToplam * rijFark) * -1;
                                        sonuc1 = (Math.Sqrt(sonuc1)) * -Convert.ToDouble(1 / θ);
                                        tbaskinlikSkoru += sonuc1;
                                    }

                                    else
                                    {
                                        sonuc1 = ((wjrToplam * rijFark) / wjr1) * -1; //karekökün içi - olamaz o yüzden -1 ile çarptım
                                        sonuc1 = (Math.Sqrt(sonuc1)) * -Convert.ToDouble(1 / θ);
                                        tbaskinlikSkoru += sonuc1;
                                    }
                                }

                            }
                        }
                    }
                    tBaskinlikSkorlari.Add(tbaskinlikSkoru);
                    tbaskinlikSkoru = 0;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Baskınlık Skorları Hesaplanamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void btnGBS_eAktar_Click(object sender, EventArgs e)
        {
            if (dataGridViewGenelSkor.Rows.Count < 150)
            {
                genelBaskinlikSkorExcelDirektAktar();
                //genelBaskinlikSkorExcelAktar();
            }
            else if (dataGridViewGenelSkor.Rows.Count >= 150)
            {
                genelBaskinlikSkorExcelDirektAktarBuyuk();
            }


        }
        public void genelBaskinlikSkorExcelDirektAktarBuyuk()
        {
            try
            {
                ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Worksheets1");
                OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                worksheet.Cells[1, 1].Value = "GENEL BASKINLIK SKORLARI";
                worksheet.Cells[1, 1, 1, 8].Merge = true;

                var columns = dataGridViewGenelSkor.Columns;
                for (int i = 0; i < columns.Count; i++)
                {
                    worksheet.Cells[2, i + 1].Value = columns[i].HeaderText;
                }

                int rowIndex = 3;
                var rows = dataGridViewGenelSkor.Rows;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[rowIndex, j + 1].Value = rows[i].Cells[j].Value;
                        }
                        rowIndex++;
                    }
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                saveFileDialog.ShowDialog();

                Stream stream = saveFileDialog.OpenFile();
                package.SaveAs(stream);
                stream.Close();

                MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void genelBaskinlikSkorExcelAktar()
        {
            try
            {
                if (dataGridViewGenelSkor.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    xcelApp.Application.Workbooks.Add(Type.Missing);

                    for (int i = 1; i < dataGridViewGenelSkor.Columns.Count + 1; i++)
                    {
                        xcelApp.Cells[1, i] = dataGridViewGenelSkor.Columns[i - 1].HeaderText;
                    }

                    for (int i = 0; i < dataGridViewGenelSkor.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewGenelSkor.Columns.Count; j++)
                        {
                            xcelApp.Cells[i + 2, j + 1] = dataGridViewGenelSkor.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Excel'e aktarılamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void genelBaskinlikSkorExcelDirektAktar()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
                saveFileDialog.FilterIndex = 3;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new ExcelFile();
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    // From DataGridView to ExcelFile.
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewGenelSkor, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

                    workbook.Save(saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Excel'e aktarma işlemi başarısız!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        private void btnKararMatEAktar_Click(object sender, EventArgs e)
        {
            if (dataGridViewKararMat.Rows.Count < 150)
            {
                kararMatDirektAktar();

            }
            else if (dataGridViewKararMat.Rows.Count >= 150)
            {
                kararMatEAktarBuyuk();
            }
        }
        public void kararMatEAktarBuyuk()
        {
            try
            {
                ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Worksheets1");
                OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                worksheet.Cells[1, 1].Value = "KARAR MATRİSİ";
                worksheet.Cells[1, 1, 1, 4].Merge = true;

                var columns = dataGridViewKararMat.Columns;
                for (int i = 0; i < columns.Count; i++)
                {
                    worksheet.Cells[2, i + 1].Value = columns[i].HeaderText;
                }

                int rowIndex = 3;
                var rows = dataGridViewKararMat.Rows;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[rowIndex, j + 1].Value = rows[i].Cells[j].Value;
                        }
                        rowIndex++;
                    }
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                saveFileDialog.ShowDialog();

                Stream stream = saveFileDialog.OpenFile();
                package.SaveAs(stream);
                stream.Close();

                MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void kararMatExcelAktar()
        {
            try
            {
                if (dataGridViewKararMat.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    xcelApp.Application.Workbooks.Add(Type.Missing);

                    for (int i = 1; i < dataGridViewKararMat.Columns.Count + 1; i++)
                    {
                        xcelApp.Cells[1, i] = dataGridViewKararMat.Columns[i - 1].HeaderText;
                    }

                    for (int i = 0; i < dataGridViewKararMat.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewKararMat.Columns.Count; j++)
                        {
                            xcelApp.Cells[i + 2, j + 1] = dataGridViewKararMat.Rows[i].Cells[j].Value.ToString();
                        }
                    }

                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;

                }
            }
            catch (Exception)
            {

                MessageBox.Show("Excel'e aktarılamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void kararMatDirektAktar()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
                saveFileDialog.FilterIndex = 3;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new ExcelFile();
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    // From DataGridView to ExcelFile.
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewKararMat, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

                    workbook.Save(saveFileDialog.FileName);


                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Excel'e aktarma işlemi başarısız!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void normalizeMatExcelAc()
        {
            try
            {
                if (dataGridViewNormalize.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    xcelApp.Application.Workbooks.Add(Type.Missing);

                    for (int i = 1; i < dataGridViewNormalize.Columns.Count + 1; i++)
                    {
                        xcelApp.Cells[1, i] = dataGridViewNormalize.Columns[i - 1].HeaderText;
                    }

                    for (int i = 0; i < dataGridViewNormalize.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewNormalize.Columns.Count; j++)
                        {
                            xcelApp.Cells[i + 2, j + 1] = dataGridViewNormalize.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Excel'e aktarılamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        private void btnNormalizeEAktar_Click(object sender, EventArgs e)
        {
            if (dataGridViewNormalize.Rows.Count < 150)
            {
                normalizeMatEAktar();
            }
            else if (dataGridViewNormalize.Rows.Count >= 150)
            {
                normlizeMatEAktarBuyuk();
            }

        }
        public void normlizeMatEAktarBuyuk()
        {
            try
            {
                ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Worksheets1");
                OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                worksheet.Cells[1, 1].Value = "NORMALİZE EDİLMİŞ KARAR MATRİSİ";
                worksheet.Cells[1, 1, 1, 4].Merge = true;

                var columns = dataGridViewNormalize.Columns;
                for (int i = 0; i < columns.Count; i++)
                {
                    worksheet.Cells[2, i + 1].Value = columns[i].HeaderText;
                }

                int rowIndex = 3;
                var rows = dataGridViewNormalize.Rows;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[rowIndex, j + 1].Value = rows[i].Cells[j].Value;
                        }
                        rowIndex++;
                    }
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                saveFileDialog.ShowDialog();

                Stream stream = saveFileDialog.OpenFile();
                package.SaveAs(stream);
                stream.Close();

                MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void normalizeMatEAktar()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
                saveFileDialog.FilterIndex = 3;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new ExcelFile();
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    // From DataGridView to ExcelFile.
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewNormalize, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

                    workbook.Save(saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Excel'e aktarma işlemi başarısız!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        private void btnBSkorEAktar_Click(object sender, EventArgs e)
        {
            if (dataGridViewBaskinlikSkoru.Rows.Count < 150)
            {
                baskinlikSkorAramaEAktar();
            }

            else if (dataGridViewBaskinlikSkoru.Rows.Count >= 150)
            {
                baskinlikSkorAramaEAktarBuyuk();
            }

        }
        public void baskinlikSkorAramaEAktarBuyuk()
        {
            try
            {
                ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Worksheets1");
                OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                worksheet.Cells[1, 1].Value = label7.Text;
                worksheet.Cells[1, 1, 1, 8].Merge = true;

                var columns = dataGridViewBaskinlikSkoru.Columns;
                for (int i = 0; i < columns.Count; i++)
                {
                    worksheet.Cells[2, i + 1].Value = columns[i].HeaderText;
                }

                int rowIndex = 3;
                var rows = dataGridViewBaskinlikSkoru.Rows;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[rowIndex, j + 1].Value = rows[i].Cells[j].Value;
                        }
                        rowIndex++;
                    }
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                saveFileDialog.ShowDialog();

                Stream stream = saveFileDialog.OpenFile();
                package.SaveAs(stream);
                stream.Close();

                MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void baskinlikSkorAramaEAktar()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
                saveFileDialog.FilterIndex = 3;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new ExcelFile();
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    // From DataGridView to ExcelFile.
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewBaskinlikSkoru, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

                    workbook.Save(saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Excel'e aktarma işlemi başarısız!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void baskinlikSkorExcelAc()
        {
            try
            {
                if (dataGridViewBaskinlikSkoru.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    xcelApp.Application.Workbooks.Add(Type.Missing);

                    for (int i = 1; i < dataGridViewBaskinlikSkoru.Columns.Count + 1; i++)
                    {
                        xcelApp.Cells[1, i] = dataGridViewBaskinlikSkoru.Columns[i - 1].HeaderText;
                    }

                    for (int i = 0; i < dataGridViewBaskinlikSkoru.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewBaskinlikSkoru.Columns.Count; j++)
                        {
                            xcelApp.Cells[i + 2, j + 1] = dataGridViewBaskinlikSkoru.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Excel'e aktarılamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        private void btnIkiliKMatEAktar_Click(object sender, EventArgs e)
        {
            if (dataGridViewKarsilastirmaMat.Rows.Count + dataGridViewKriterAgirliklari.Rows.Count < 150)
            {
                ikiliKMatExcelDirektAktar();
            }
            else if (dataGridViewKarsilastirmaMat.Rows.Count + dataGridViewKriterAgirliklari.Rows.Count + 10 >= 150)
            {
                ikiliKMatExcelDirektAktarBuyuk();
            }
        }
        public void ikiliKMatExcelDirektAktarBuyuk()
        {
            try
            {
                ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Worksheets1");
                OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                worksheet.Cells[1, 1].Value = "KRİTERLER İÇİN İKİLİ KARŞILAŞTIRMA MATRİSİ";
                worksheet.Cells[1, 1, 1, 6].Merge = true;
                int satir = 2;
                var columns = dataGridViewKarsilastirmaMat.Columns;
                for (int i = 0; i < columns.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = columns[i].HeaderText;
                }

                satir++;
                var rows = dataGridViewKarsilastirmaMat.Rows;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = rows[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }

                satir += 3;

                worksheet.Cells[satir, 1].Value = "KRİTERLER İÇİN İKİLİ KARŞILAŞTIRMA MATRİSİ";
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;
                var sutunlar = dataGridViewKriterAgirliklari.Columns;
                for (int i = 0; i < sutunlar.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = sutunlar[i].HeaderText;
                }

                satir++;
                var satirlar = dataGridViewKarsilastirmaMat.Rows;
                for (int i = 0; i < satirlar.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < satirlar[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = satirlar[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }


                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                saveFileDialog.ShowDialog();

                Stream stream = saveFileDialog.OpenFile();
                package.SaveAs(stream);
                stream.Close();

                MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void ikiliKMatExcelDirektAktar()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
                saveFileDialog.FilterIndex = 3;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new ExcelFile();
                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    int say = 0;
                    worksheet.Cells[0, 0].Value = "KRİTERLER İÇİN İKİLİ KARŞILAŞTIRMA MATRİSİ: ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewKarsilastirmaMat, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say = dataGridViewKarsilastirmaMat.Rows.Count + 4;
                    worksheet.Cells[say, 0].Value = "KRİTER AĞIRLIKLARI: ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewKriterAgirliklari, new ImportFromDataGridViewOptions()
                    {
                        ColumnHeaders = true,
                        StartRow = say
                    });
                    workbook.Save(saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Excel'e aktarma işlemi başarısız!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public void ikiliKMatExcelAc()
        {
            try
            {
                if (dataGridViewKarsilastirmaMat.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    xcelApp.Application.Workbooks.Add(Type.Missing);

                    for (int i = 1; i < dataGridViewKarsilastirmaMat.Columns.Count + 1; i++)
                    {
                        xcelApp.Cells[1, i] = dataGridViewKarsilastirmaMat.Columns[i - 1].HeaderText;
                    }

                    for (int i = 0; i < dataGridViewKarsilastirmaMat.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewKarsilastirmaMat.Columns.Count; j++)
                        {
                            xcelApp.Cells[i + 2, j + 1] = dataGridViewKarsilastirmaMat.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    if (dataGridViewKriterAgirliklari.Rows.Count > 0)
                    {

                        for (int i = 1; i < dataGridViewKriterAgirliklari.Columns.Count + 1; i++)
                        {
                            xcelApp.Cells[dataGridViewKarsilastirmaMat.Rows.Count + 4, i] = dataGridViewKriterAgirliklari.Columns[i - 1].HeaderText;
                        }
                        int a = 0;
                        for (int i = dataGridViewKarsilastirmaMat.Rows.Count; i < dataGridViewKarsilastirmaMat.Rows.Count + dataGridViewKriterAgirliklari.Rows.Count; i++)
                        {
                            int s = 0;
                            for (int j = 0; j < dataGridViewKriterAgirliklari.Columns.Count; j++)
                            {
                                xcelApp.Cells[i + 5, j + 1] = dataGridViewKriterAgirliklari.Rows[a].Cells[s].Value.ToString();
                                s++;
                            }
                            a++;
                        }
                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarılamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void btnKriterAgirlikKaydet_Click(object sender, EventArgs e)
        {
            sonuçlarToolStripMenuItem.Visible = true;
            kriterAgirlikKaydet();
            tabControl1.SelectedTab = tabPageGenelBaskinlik;
        }
        public void kriterAgirlikKaydet()
        {
            agirliklar.Clear();
            try
            {
                for (int j = 1; j < kriterler.Count + 1; j++)
                {
                    agirliklar.Add(Convert.ToDouble(dataGridViewKriterAgirliklari.Rows[0].Cells[j].Value));
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            goreliAgirliklarim();
            maxGenelBakinlik();
            genelBaskinlikSkoruMatrisi();
        }
        private void btnAgirlikDuzenle_Click(object sender, EventArgs e)
        {
            panel82.Visible = false;
            panel83.Visible = false;
            panel80.Visible = false;
            label24.Visible = false;
            lblTutOrani.Visible = false;
            btnKriterAgirlikKaydet.Visible = false;
            lblUyari.Visible = false;
            panel52.Visible = false;
            btnKriterAgirlikKaydet.Text = "AĞIRLIKLARI KAYDET";
            tabControl1.SelectedTab = tabPageKarsilastirmaMat;
        }
        public void baskinlikSkoruMaxMin() //tbaskınlık skorları listesindeki max ve min baskınlık skorlarını bulan metod
        {
            tümBaskinlikSkorlari();
            //karar matrisindeki bir sutundaki max ve min değerleri bulan döngüler
            maxBaskinlikSkoru = Convert.ToDouble(tBaskinlikSkorlari[0]);
            minBaskinlikSkoru = Convert.ToDouble(tBaskinlikSkorlari[0]);
            for (int i = 0; i < tBaskinlikSkorlari.Count; i++)
            {
                baskinlikSkorum = Convert.ToDouble(tBaskinlikSkorlari[i]);
                if (baskinlikSkorum > maxBaskinlikSkoru)
                {
                    maxBaskinlikSkoru = baskinlikSkorum;
                }
            }
            for (int i = 0; i < tBaskinlikSkorlari.Count; i++)
            {
                baskinlikSkorum = Convert.ToDouble(tBaskinlikSkorlari[i]);
                if (baskinlikSkorum < minBaskinlikSkoru)
                {
                    minBaskinlikSkoru = baskinlikSkorum;
                }
            }
        }
        public void genelBaskinlik()
        {
            baskinlikSkoruMaxMin();
            genelBaskinlikSkorlari.Clear(); //teta için ekledim
            for (int i = 0; i < tBaskinlikSkorlari.Count; i++)
            {
                bskor = Convert.ToDouble(tBaskinlikSkorlari[i]);
                genelBaskinlikSkoru = ((bskor - minBaskinlikSkoru) / (maxBaskinlikSkoru - minBaskinlikSkoru));
                genelBaskinlikSkorlari.Add(genelBaskinlikSkoru);
                genelBaskinlikSkoru = 0;
            }
        }
        public void maxGenelBakinlik()
        {
            genelBaskinlik();
            int genelBakinlikSira = 0;
            maxGenelBaskinlikSkor = Convert.ToDouble(genelBaskinlikSkorlari[0]);
            for (int i = 0; i < genelBaskinlikSkorlari.Count; i++)
            {
                gSkor = Convert.ToDouble(genelBaskinlikSkorlari[i]);
                if (gSkor > maxGenelBaskinlikSkor)
                {
                    maxGenelBaskinlikSkor = gSkor;
                    genelBakinlikSira = i;
                }
            }
            enİyiAlternatif = alternatifler[genelBakinlikSira].ToString();
            lblEnİyiAlternatif.Text = enİyiAlternatif;
        }
        public void genelBaskinlikSkoruMatrisi()
        {
            dataGridViewGenelSkor.Rows.Clear();
            dataGridViewGenelSkor.ColumnCount = 3;
            dataGridViewGenelSkor.Columns[0].Name = " ";
            dataGridViewGenelSkor.Columns[1].Name = "S(Ai,Ai')";
            dataGridViewGenelSkor.Columns[2].Name = "Si";
            for (int i = 0; i < alternatifler.Count; i++)
            {
                dataGridViewGenelSkor.Rows.Add(alternatifler[i].ToString());
                dataGridViewGenelSkor.Rows[i].Cells[1].Value = tBaskinlikSkorlari[i].ToString();
                dataGridViewGenelSkor.Rows[i].Cells[2].Value = genelBaskinlikSkorlari[i].ToString();
            }
            //ilk sutundaki değerlerin değiştirilmesini önleme
            for (int i = 0; i < alternatifler.Count; i++)
            {
                dataGridViewGenelSkor.Rows[i].Cells[0].ReadOnly = true;
            }

        }
        protected void RadioChange(object sender, EventArgs e)
        {

        }
        private void btnKarsilastirmaMatOlustur_Click(object sender, EventArgs e)
        {

            if (panel80.Visible == true)
            {
                panel80.Visible = false;
                lblTutOrani.Visible = false;
                lblUyari.Visible = false;
                label24.Visible = false;
                btnKriterAgirlikKaydet.Visible = false;
                btnAhpAyrintiCozGoster.Visible = false;
                panel52.Visible = false;
            }
            if (tabControl1.SelectedTab == tabPageKarsilastirmaMat)
            {
                panel80.Visible = false;
                lblTutOrani.Visible = false;
                lblUyari.Visible = false;
                label24.Visible = false;
                btnKriterAgirlikKaydet.Visible = false;
                btnAhpAyrintiCozGoster.Visible = false;
                panel52.Visible = false;
            }
            agirlikDegerleriToolStripMenuItem.Visible = true;
            try
            {
                karsilastirmaMatOlustur();
                for (int i = 0; i < dataGridViewKarsilastirmaMat.Rows.Count; i++)
                {
                    for (int j = 1; j < dataGridViewKarsilastirmaMat.Columns.Count; j++)
                    {
                        if (dataGridViewKarsilastirmaMat.Rows[i].Cells[j].Value == null)
                        {
                            MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                wVektörü();
                ayrintiKMatDoldur();
                DVektörü();
                tutarlilikOrani();
                tabControl1.SelectedTab = tabPageAHP;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void karsilastirmaMatOlustur() //dinamik oluşturduğum buton
        {
            try
            {
                btnAhpAyrintiCozGoster.Visible = false;

                for (int i = 1; i < kriterler.Count; i++)
                {
                    for (int j = i; j < kriterler.Count; j++)
                    {
                        int J = j + 1;
                        for (int rb = 9; rb > 0; rb--)
                        {
                            if (radioButton[y].Checked == true)
                            {
                                if (i == 1)
                                {

                                }
                                dataGridViewKarsilastirmaMat.Rows[i - 1].Cells[J].Value = Convert.ToDouble(rb);
                                dataGridViewKarsilastirmaMat.Rows[j].Cells[i].Value = Convert.ToDouble(1 / Convert.ToDouble(rb));
                            }

                            y++;
                        }

                        for (int rb = 2; rb < 10; rb++)
                        {
                            if (radioButton1[x].Checked == true)
                            {
                                dataGridViewKarsilastirmaMat.Rows[j].Cells[i].Value = Convert.ToDouble(rb);
                                dataGridViewKarsilastirmaMat.Rows[i - 1].Cells[J].Value = Convert.ToDouble(1 / Convert.ToDouble(rb));
                            }
                            x++;
                        }
                    }

                }
                y = 0;
                x = 0;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        //kriterlerin birbirlerine göre önem değerlerini gösterir
        public void cMatrisi()
        {
            try
            {
                cMatrisTasarim();
                double pay, sonuc;
                paydaHesapla();
                for (int j = 1; j < kriterler.Count + 1; j++)
                {
                    for (int i = 0; i < kriterler.Count; i++)
                    {
                        pay = Convert.ToDouble(dataGridViewKarsilastirmaMat.Rows[i].Cells[j].Value);
                        sonuc = pay / Convert.ToDouble(paydaListesi[j - 1]);
                        dataGridViewC.Rows[i].Cells[j].Value = sonuc;
                        sonuc = 0;
                        pay = 0;
                    }

                }
            }
            catch (Exception)
            {

                MessageBox.Show("C matrisi oluşturulamadı.!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void cMatrisTasarim()
        {
            dataGridViewC.Rows.Clear();
            dataGridViewC.Columns.Clear();

            int k = 1;
            dataGridViewC.ColumnCount = kriterler.Count + 1;
            dataGridViewC.Columns[0].Name = " ";

            for (int i = 0; i < kriterler.Count; i++)
            {
                k = 1;
                for (int j = 0; j < kriterler.Count; j++)
                {
                    dataGridViewC.Columns[k].Name = kriterler[j].ToString();
                    k++;
                }
            }
            for (int j = 0; j < kriterler.Count; j++)
            {
                dataGridViewC.Rows.Add(kriterler[j].ToString());
            }
            for (int cR = 0; cR < kriterler.Count; cR++)
            {
                dataGridViewC.Rows[cR].Cells[0].ReadOnly = true;
            }
        }
        private void btnSonucExcelAktar_Click(object sender, EventArgs e)
        {
            int say = 0;
            say = dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count + dataGridViewSonucKarsilastirmaMat.Rows.Count + dataGridViewSonucGenelBaskinlik.Rows.Count + 30;
            if (say < 150)
            {
                sonucExcelDirektAktar();
                //sonucExcelAktar();
            }
            if (say >= 150)
            {
                sonucExcelAktarBuyuk();
            }

        }
        public void sonucExcelAktarBuyuk()
        {
            try
            {
                ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Worksheets1");
                OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                worksheet.Cells[1, 1].Value = "KARAR MATRİSİ";
                worksheet.Cells[1, 1, 1, 6].Merge = true;
                int satir = 2;
                var columns = dataGridViewSonucKararMat.Columns;
                for (int i = 0; i < columns.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = columns[i].HeaderText;
                }

                satir++;
                var rows = dataGridViewSonucKararMat.Rows;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = rows[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }



                satir += 3;

                worksheet.Cells[satir, 1].Value = "NORMALİZE EDİLMİŞ KARAR MATRİSİ";
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;
                var sutunlar = dataGridViewSonucNormalizeMat.Columns;
                for (int i = 0; i < sutunlar.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = sutunlar[i].HeaderText;
                }

                satir++;
                var satirlar = dataGridViewSonucNormalizeMat.Rows;
                for (int i = 0; i < satirlar.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < satirlar[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = satirlar[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }


                satir += 3;

                worksheet.Cells[satir, 1].Value = "KRİTER AĞIRLIKLARI";
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;

                for (int i = 0; i < dataGridViewSonucKriterAgirlik.Columns.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = dataGridViewSonucKriterAgirlik.Columns[i].HeaderText;
                }

                satir++;

                for (int i = 0; i < dataGridViewSonucKriterAgirlik.Rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < dataGridViewSonucKriterAgirlik.Rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = dataGridViewSonucKriterAgirlik.Rows[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }




                satir += 3;

                worksheet.Cells[satir, 1].Value = "GÖRELİ AĞIRLIKLAR";
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;

                for (int i = 0; i < dataGridViewSonucGoreliAgirlik.Columns.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = dataGridViewSonucGoreliAgirlik.Columns[i].HeaderText;
                }

                satir++;

                for (int i = 0; i < dataGridViewSonucGoreliAgirlik.Rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < dataGridViewSonucGoreliAgirlik.Rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = dataGridViewSonucGoreliAgirlik.Rows[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }


                if (yontem == "ahp")
                {
                    satir += 3;

                    worksheet.Cells[satir, 1].Value = "İKİLİ KARŞILAŞTIRMA MATRİSİ";
                    worksheet.Cells[satir, 1, satir, 6].Merge = true;
                    satir++;

                    for (int i = 0; i < dataGridViewSonucKarsilastirmaMat.Columns.Count; i++)
                    {
                        worksheet.Cells[satir, i + 1].Value = dataGridViewSonucKarsilastirmaMat.Columns[i].HeaderText;
                    }

                    satir++;

                    for (int i = 0; i < dataGridViewSonucKarsilastirmaMat.Rows.Count; i++)
                    {
                        if (rows[i].Cells[0] != null)
                        {
                            for (int j = 0; j < dataGridViewSonucKarsilastirmaMat.Rows[i].Cells.Count; j++)
                            {
                                worksheet.Cells[satir, j + 1].Value = dataGridViewSonucKarsilastirmaMat.Rows[i].Cells[j].Value;
                            }
                            satir++;
                        }
                    }


                }



                satir += 3;

                worksheet.Cells[satir, 1].Value = "GENEL BASKINLIK SKORLARI";
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;

                for (int i = 0; i < dataGridViewSonucGenelBaskinlik.Columns.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = dataGridViewSonucGenelBaskinlik.Columns[i].HeaderText;
                }

                satir++;

                for (int i = 0; i < dataGridViewSonucGenelBaskinlik.Rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < dataGridViewSonucGenelBaskinlik.Rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = dataGridViewSonucGenelBaskinlik.Rows[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }




                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                saveFileDialog.ShowDialog();

                Stream stream = saveFileDialog.OpenFile();
                package.SaveAs(stream);
                stream.Close();

                MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void sonucExcelDirektAktar()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
                saveFileDialog.FilterIndex = 3;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new ExcelFile();
                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    int say = 0;
                    worksheet.Cells[0, 0].Value = "KARAR MATRİSİ: ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewSonucKararMat, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewSonucKararMat.Rows.Count + 3;
                    worksheet.Cells[say, 0].Value = "NORMALİZE EDİLMİŞ KARAR MATRİSİ: ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewSonucNormalizeMat, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewSonucNormalizeMat.Rows.Count + 3;
                    worksheet.Cells[say, 0].Value = "KRİTER AĞIRLIKLARI: ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewSonucKriterAgirlik, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewSonucKriterAgirlik.Rows.Count + 3;
                    worksheet.Cells[say, 0].Value = "GÖRELİ AĞIRLIKLAR: ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewSonucGoreliAgirlik, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewSonucGoreliAgirlik.Rows.Count + 3;
                    worksheet.Cells[say, 0].Value = "İKİLİ KARŞILAŞTIRMA MATRİSİ: ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewSonucKarsilastirmaMat, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewSonucKarsilastirmaMat.Rows.Count + 3;
                    worksheet.Cells[say, 0].Value = "GENEL BASKINLIK SKORLARI: ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewSonucGenelBaskinlik, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewSonucGenelBaskinlik.Rows.Count + 3;
                    workbook.Save(saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Excel'e aktarma işlemi başarısız!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public void paydaHesapla() //yüzde önem dağılımlarını hesaplamak için gereken sutun toplamlarını hesaplayıp paydaListesi' ne ekleyen metod
        {
            try
            {
                for (int j = 1; j < kriterler.Count + 1; j++)
                {
                    double payda = 0;
                    for (int i = 0; i < kriterler.Count; i++)
                    {
                        payda += Convert.ToDouble(dataGridViewKarsilastirmaMat.Rows[i].Cells[j].Value);
                    }
                    paydaListesi.Add(payda);

                    payda = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void wVektörüTasarim() //öncelik (w) vektörünü hesaplayan metod
        {
            dataGridViewWVektörü.Rows.Clear();
            dataGridViewWVektörü.Columns.Clear();
            int k = 1;
            dataGridViewWVektörü.ColumnCount = kriterler.Count + 1;
            dataGridViewWVektörü.Columns[0].Name = " ";

            for (int j = 0; j < kriterler.Count; j++)
            {
                dataGridViewWVektörü.Columns[k].Name = kriterler[j].ToString();
                k++;
            }
            for (int j = 0; j < 1; j++)
            {
                dataGridViewWVektörü.Rows.Add("Ağırlıklar");
            }
            dataGridViewWVektörü.Rows[0].Cells[0].ReadOnly = true;
        }
        public void wVektörü() //öncelik (w) vektörünü hesaplayan metod
        {
            try
            {
                cMatrisi();
                wVektörüTasarim();
                for (int i = 0; i < kriterler.Count; i++)
                {
                    double satirToplam = 0;
                    for (int j = 1; j < kriterler.Count + 1; j++)
                    {
                        satirToplam += Convert.ToDouble(dataGridViewC.Rows[i].Cells[j].Value);
                    }
                    dataGridViewWVektörü.Rows[0].Cells[i + 1].Value = (satirToplam / kriterler.Count);
                    agirliklar.Add(satirToplam / kriterler.Count);
                    satirToplam = 0;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("V vektörü oluşturulamadı!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public void ayrintiKMatDoldur()
        {
            try
            {
                dataGridViewAyrintiKmat.Columns.Clear();
                dataGridViewAyrintiKmat.Rows.Clear();
                int k = 1;
                dataGridViewAyrintiKmat.ColumnCount = dataGridViewKarsilastirmaMat.Columns.Count;
                for (int i = 0; i < kriterler.Count; i++)
                {
                    k = 1;
                    for (int j = 0; j < kriterler.Count; j++)
                    {
                        dataGridViewAyrintiKmat.Columns[k].Name = kriterler[j].ToString();
                        k++;
                    }
                }
                for (int j = 0; j < kriterler.Count; j++)
                {
                    dataGridViewAyrintiKmat.Rows.Add(kriterler[j].ToString());
                }
                for (int cR = 0; cR < kriterler.Count; cR++)
                {
                    dataGridViewAyrintiKmat.Rows[cR].Cells[0].ReadOnly = true;
                }
                for (int i = 0; i < dataGridViewKarsilastirmaMat.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridViewKarsilastirmaMat.Columns.Count; j++)
                    {
                        dataGridViewAyrintiKmat.Rows[i].Cells[j].Value = dataGridViewKarsilastirmaMat.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public void DVektörüTasarim() //öncelik (w) vektörünü hesaplayan metod
        {
            dataGridViewDVektör.Columns.Clear();
            dataGridViewDVektör.Rows.Clear();
            int k = 1;
            dataGridViewDVektör.ColumnCount = kriterler.Count + 1;
            dataGridViewDVektör.Columns[0].Name = " ";

            for (int j = 0; j < kriterler.Count; j++)
            {
                dataGridViewDVektör.Columns[k].Name = kriterler[j].ToString();
                k++;
            }
            for (int j = 0; j < 1; j++)
            {
                dataGridViewDVektör.Rows.Add("D satır vektörü");
            }

            dataGridViewDVektör.Rows[0].Cells[0].ReadOnly = true;
        }
        public void DVektörü()
        {
            try
            {
                DVektörüTasarim();
                for (int i = 0; i < kriterler.Count; i++)
                {
                    double carpim, toplam = 0;
                    for (int j = 1; j < kriterler.Count + 1; j++)
                    {
                        carpim = Convert.ToDouble(dataGridViewAyrintiKmat.Rows[i].Cells[j].Value) * Convert.ToDouble(dataGridViewWVektörü.Rows[0].Cells[j].Value);
                        toplam += carpim;
                    }
                    dataGridViewDVektör.Rows[0].Cells[i + 1].Value = toplam;
                    carpim = 0;
                    toplam = 0;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("D vektörü oluşturulamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void eToplamı()
        {
            double e, eToplam = 0;
            for (int j = 1; j < kriterler.Count + 1; j++)
            {
                e = Convert.ToDouble(dataGridViewDVektör.Rows[0].Cells[j].Value) / Convert.ToDouble(dataGridViewWVektörü.Rows[0].Cells[j].Value);
                eToplam += e;
                e = 0;
            }
            lamda = eToplam / kriterler.Count;
        }
        public void tutarlilikOrani()
        {
            try
            {
                double RI = 0;
                eToplamı();
                CI = (lamda - kriterler.Count) / (kriterler.Count - 1);
                lamda = 0;
                if (kriterler.Count == 1)
                {
                    RI = 0;
                }
                else if (kriterler.Count == 2)
                {
                    RI = 0;
                }
                else if (kriterler.Count == 3)
                {
                    RI = 0.58;
                }
                else if (kriterler.Count == 4)
                {
                    RI = 0.90;
                }
                else if (kriterler.Count == 5)
                {
                    RI = 1.12;
                }
                else if (kriterler.Count == 6)
                {
                    RI = 1.24;
                }
                else if (kriterler.Count == 7)
                {
                    RI = 1.32;
                }
                else if (kriterler.Count == 8)
                {
                    RI = 1.41;
                }
                else if (kriterler.Count == 9)
                {
                    RI = 1.45;
                }
                else if (kriterler.Count == 10)
                {
                    RI = 1.49;
                }
                else if (kriterler.Count == 11)
                {
                    RI = 1.51;
                }
                else if (kriterler.Count == 12)
                {
                    RI = 1.48;
                }
                else if (kriterler.Count == 13)
                {
                    RI = 1.56;
                }
                else if (kriterler.Count == 14)
                {
                    RI = 1.57;
                }
                else if (kriterler.Count == 15)
                {
                    RI = 1.59;
                }
                else if (kriterler.Count >= 16)
                {
                    MessageBox.Show("En fazla 15 kriter için hesaplama yapılmaktadır.!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                CR = Convert.ToDouble(CI / RI);
                lblCI.Text = CI.ToString();
                lblRI.Text = RI.ToString();
                lblTutarlilikOrani.Text = "";
                lblTutarlilikOrani.Text = CR.ToString();
                lblTutOrani.Text = "";
                lblTutOrani.Text = CR.ToString();
                CR = 0;
                CI = 0;
                RI = 0;
            }
            catch (Exception)
            {
                MessageBox.Show("Tutarlılık oranı hesaplanamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void btnAyrintiliCozum_Click(object sender, EventArgs e)
        {
            kismiBaskinlikSkorlariToolStripMenuItem.Visible = true;
            try
            {
                tabControl1.SelectedTab = tabPageSonuc;
                sonucKararMatDoldur();
                sonucNormalizeMatDoldur();
                sonucKriterAgirlikDoldur();
                sonucGoreliAgirlikDoldur();
                sonucGenelBaskinlikDoldur();
                sonucKarsilastirmaMatDoldur();

            }
            catch (Exception)
            {

                MessageBox.Show("Ayrıntılı çözümler görüntülenemedi!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void sonucKararMatDoldur()
        {
            try
            {
                if (dataGridViewKararMat.Rows.Count > 0)
                {
                    dataGridViewSonucKararMat.Columns.Clear();
                    dataGridViewSonucKararMat.Rows.Clear();
                    dataGridViewSonucKararMat.ColumnCount = dataGridViewKararMat.Columns.Count;
                    dataGridViewSonucKararMat.Columns[0].Name = " ";
                    int k = 1;
                    for (int i = 0; i < kriterler.Count; i++)
                    {
                        dataGridViewSonucKararMat.Columns[k].Name = kriterler[i].ToString();
                        k++;
                    }

                    for (int i = 0; i < alternatifler.Count; i++)
                    {

                        dataGridViewSonucKararMat.Rows.Add(alternatifler[i].ToString());
                    }

                    for (int i = 0; i < dataGridViewKararMat.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewKararMat.Columns.Count; j++)
                        {

                            dataGridViewSonucKararMat.Rows[i].Cells[j].Value = dataGridViewKararMat.Rows[i].Cells[j].Value.ToString();


                        }
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Hata!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public void sonucNormalizeMatDoldur()
        {
            try
            {
                dataGridViewSonucNormalizeMat.Columns.Clear();
                dataGridViewSonucNormalizeMat.Rows.Clear();
                //normalize matrisi
                dataGridViewSonucNormalizeMat.ColumnCount = dataGridViewNormalize.Columns.Count; //alternatif sutunuda bulunması gerektiğinden kriterler+1 tane sutun ekledim

                int s = 1;
                for (int i = 0; i < kriterler.Count; i++)
                {
                    dataGridViewSonucNormalizeMat.Columns[s].Name = kriterler[i].ToString();
                    s++;
                }
                for (int i = 0; i < alternatifler.Count; i++)
                {
                    dataGridViewSonucNormalizeMat.Rows.Add(alternatifler[i].ToString());
                }
                //İLK SUTUNDAKİ DEĞERLERİN DEĞİŞTİRİLMESİNİ ENGELLEDİM
                for (int rC = 0; rC < alternatifler.Count; rC++)
                {
                    dataGridViewSonucNormalizeMat.Rows[rC].Cells[0].ReadOnly = true;
                }
                for (int i = 0; i < dataGridViewNormalize.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridViewNormalize.Columns.Count; j++)
                    {
                        dataGridViewSonucNormalizeMat.Rows[i].Cells[j].Value = dataGridViewNormalize.Rows[i].Cells[j].Value.ToString();
                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("Hata!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


        }
        public void sonucKriterAgirlikDoldur()
        {
            try
            {
                if (dataGridViewWVektörü.Rows.Count > 0)
                {
                    dataGridViewSonucKriterAgirlik.Rows.Clear();
                    dataGridViewSonucKriterAgirlik.Columns.Clear();
                    dataGridViewSonucKriterAgirlik.ColumnCount = dataGridViewWVektörü.Columns.Count;
                    int k = 1;
                    for (int j = 0; j < kriterler.Count; j++)
                    {
                        dataGridViewSonucKriterAgirlik.Columns[k].Name = kriterler[j].ToString();
                        k++;
                    }
                    for (int j = 0; j < 1; j++)
                    {

                        dataGridViewSonucKriterAgirlik.Rows.Add("Wj");
                    }


                    for (int i = 0; i < dataGridViewWVektörü.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewWVektörü.Columns.Count; j++)
                        {
                            dataGridViewSonucKriterAgirlik.Rows[i].Cells[j].Value = dataGridViewWVektörü.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }

                if (dataGridViewAgirlik.Rows.Count > 0)
                {
                    dataGridViewSonucKriterAgirlik.Rows.Clear();
                    dataGridViewSonucKriterAgirlik.Columns.Clear();
                    dataGridViewSonucKriterAgirlik.ColumnCount = dataGridViewAgirlik.Columns.Count;
                    int k = 1;
                    for (int j = 0; j < kriterler.Count; j++)
                    {
                        dataGridViewSonucKriterAgirlik.Columns[k].Name = kriterler[j].ToString();
                        k++;
                    }
                    for (int j = 0; j < 1; j++)
                    {

                        dataGridViewSonucKriterAgirlik.Rows.Add("Ağırlıklar");
                    }


                    for (int i = 0; i < dataGridViewAgirlik.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewAgirlik.Columns.Count; j++)
                        {
                            dataGridViewSonucKriterAgirlik.Rows[i].Cells[j].Value = dataGridViewAgirlik.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        public void sonucGoreliAgirlikDoldur()
        {
            try
            {
                dataGridViewSonucGoreliAgirlik.Rows.Clear();
                dataGridViewSonucGoreliAgirlik.Columns.Clear();

                dataGridViewSonucGoreliAgirlik.ColumnCount = kriterler.Count + 2;
                dataGridViewSonucGoreliAgirlik.Columns[0].Name = " ";
                dataGridViewSonucGoreliAgirlik.Columns[kriterler.Count + 1].Name = "Toplam";
                int k = 1;
                for (int j = 0; j < kriterler.Count; j++)
                {
                    dataGridViewSonucGoreliAgirlik.Columns[k].Name = kriterler[j].ToString();
                    k++;
                }

                dataGridViewSonucGoreliAgirlik.Rows.Add("Wjr");
                dataGridViewSonucGoreliAgirlik.Rows[0].Cells[0].ReadOnly = true;

                int J = 1;
                for (int i = 0; i < goreliAgirliklar.Count; i++)
                {
                    dataGridViewSonucGoreliAgirlik.Rows[0].Cells[J].Value = goreliAgirliklar[i].ToString();
                    J++;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        public void sonucGenelBaskinlikDoldur()
        {
            try
            {
                dataGridViewSonucGenelBaskinlik.Columns.Clear();
                dataGridViewSonucGenelBaskinlik.Rows.Clear();
                dataGridViewSonucGenelBaskinlik.ColumnCount = 3;
                dataGridViewSonucGenelBaskinlik.Columns[0].Name = " ";
                dataGridViewSonucGenelBaskinlik.Columns[1].Name = "S(Ai,Ai')";
                dataGridViewSonucGenelBaskinlik.Columns[2].Name = "Si";

                for (int i = 0; i < alternatifler.Count; i++)
                {
                    dataGridViewSonucGenelBaskinlik.Rows.Add(alternatifler[i].ToString());
                    dataGridViewSonucGenelBaskinlik.Rows[i].Cells[1].Value = tBaskinlikSkorlari[i].ToString();
                    dataGridViewSonucGenelBaskinlik.Rows[i].Cells[2].Value = genelBaskinlikSkorlari[i].ToString();
                }
                dataGridViewSonucGenelBaskinlik.Rows[0].Cells[0].ReadOnly = true;
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void btnAhpAyrintiCozGoster_Click(object sender, EventArgs e)
        {
            ahpAyrintiliCozumToolStripMenuItem.Visible = true;
            //tutarlılık oranı hesaplanırken zaten gridler doldurulmuştu.
            tabControl1.SelectedTab = tabPageAhpAyrinti;
        }
        private void btnAhpAyrintiAgirlikKaydet_Click(object sender, EventArgs e)
        {
            sonuçlarToolStripMenuItem.Visible = true;
            kriterAgirlikKaydet();
            tabControl1.SelectedTab = tabPageGenelBaskinlik;
        }
        private void btnKararMatNormalize2_Click(object sender, EventArgs e)
        {

            secilenNormalizeYontemi = "2";
            normalizeEdilmisKararMatrisiToolStripMenuItem.Visible = true;
            normalizeMatCerceve();
            for (int i = 0; i < alternatifler.Count; i++) //satır
            {
                for (int j = 1; j < kriterler.Count + 1; j++) //sutun
                {
                    if (dataGridViewKararMat.Rows[i].Cells[j].Value == null)
                    {
                        MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
            }
            normalizeYeni();

            if (tiklanma == "eski")
            {
                pnlYontemSec.Visible = true;
            }
        }
        private void btnAhpKriterAgrBulEAktar_Click(object sender, EventArgs e)
        {
            //int satirSay = 0;
            //satirSay = 30 + dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + dataGridViewDVektör.Rows.Count;
            //if (satirSay < 150)
            //{
            //    ahpAyrintiliCozumExcelDirektAktar();
            //    //ahpAyrintiliCozumExcelAktar();
            //}
            //else if (satirSay > 150)
            //{
            //    ahpAyrintiliCozumEAktarBuyuk();
            //}

            ahpAyrintiliCozumEAktarBuyuk();

        }
        public void ahpAyrintiliCozumEAktarBuyuk()
        {
            try
            {
                ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Worksheets1");
                OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                worksheet.Cells[1, 1].Value = "KRİTERLER İÇİN İKİLİ KARŞILAŞTIRMA MATRİSİ";
                worksheet.Cells[1, 1, 1, 6].Merge = true;
                int satir = 2;
                var columns = dataGridViewAyrintiKmat.Columns;
                for (int i = 0; i < columns.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = columns[i].HeaderText;
                }

                satir++;
                var rows = dataGridViewAyrintiKmat.Rows;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = rows[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }

                satir += 3;

                worksheet.Cells[satir, 1].Value = "C MATRİSİ";
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;
                var sutunlar = dataGridViewC.Columns;
                for (int i = 0; i < sutunlar.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = sutunlar[i].HeaderText;
                }

                satir++;
                var satirlar = dataGridViewC.Rows;
                for (int i = 0; i < satirlar.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < satirlar[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = satirlar[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }


                satir += 3;

                worksheet.Cells[satir, 1].Value = "W VEKTÖRÜ";
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;

                for (int i = 0; i < dataGridViewWVektörü.Columns.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = dataGridViewWVektörü.Columns[i].HeaderText;
                }

                satir++;

                for (int i = 0; i < dataGridViewWVektörü.Rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < dataGridViewWVektörü.Rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = dataGridViewWVektörü.Rows[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }




                satir += 3;

                worksheet.Cells[satir, 1].Value = "D VEKTÖRÜ";
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;

                for (int i = 0; i < dataGridViewDVektör.Columns.Count; i++)
                {
                    worksheet.Cells[satir, i + 1].Value = dataGridViewDVektör.Columns[i].HeaderText;
                }

                satir++;

                for (int i = 0; i < dataGridViewDVektör.Rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < dataGridViewDVektör.Rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[satir, j + 1].Value = dataGridViewDVektör.Rows[i].Cells[j].Value;
                        }
                        satir++;
                    }
                }
                satir += 3;

                worksheet.Cells[satir, 1].Value = "CI:   " + lblCI.Text;
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;
                worksheet.Cells[satir, 1].Value = "RI:   " + lblRI.Text;
                worksheet.Cells[satir, 1, satir, 6].Merge = true;
                satir++;
                worksheet.Cells[satir, 1].Value = "Tutarlılık Oranı:   " + lblTutarlilikOrani.Text;
                worksheet.Cells[satir, 1, satir, 6].Merge = true;


                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                saveFileDialog.ShowDialog();

                Stream stream = saveFileDialog.OpenFile();
                package.SaveAs(stream);
                stream.Close();

                MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void ahpAyrintiliCozumExcelDirektAktar()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
                saveFileDialog.FilterIndex = 3;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new ExcelFile();
                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    int say = 0;
                    worksheet.Cells[0, 0].Value = "KRİTERLER İÇİN İKİLİ KARŞILAŞTIRMA MATRİSİ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewAyrintiKmat, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewAyrintiKmat.Rows.Count + 3;
                    worksheet.Cells[say, 0].Value = "C MATRİSİ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewC, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewC.Rows.Count + 3;
                    worksheet.Cells[say, 0].Value = "AĞIRLIKLAR";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewWVektörü, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewWVektörü.Rows.Count + 3;
                    worksheet.Cells[say, 0].Value = "D MATRİSİ";
                    say += 2;
                    DataGridViewConverter.ImportFromDataGridView(worksheet, this.dataGridViewDVektör, new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
                    say += dataGridViewDVektör.Rows.Count + 3;
                    worksheet.Cells[say, 0].Value = "CI: " + lblCI.Text;
                    say += 2;
                    worksheet.Cells[say, 0].Value = "RI: " + lblRI.Text;
                    say += 2;
                    worksheet.Cells[say, 0].Value = "TUTARLILIK ORANI: " + lblTutOrani.Text;


                    workbook.Save(saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Excel'e aktarma işlemi başarısız!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void btnYenile_Click(object sender, EventArgs e)
        {
            tumunuTemizle();
            tiklanma = "olustur";
        }
        public void tumunuTemizle()
        {
            try
            {
                tiklanma = "";
                calismaIsmi = "";
                pnlYontemSec.Visible = true;
                panel12.Visible = false;

                btnKriterAgirlikKaydet.Visible = false;
                btnAhpAyrintiCozGoster.Visible = false;
                panel82.Visible = false;
                panel83.Visible = false;
                panel80.Visible = false;
                label24.Visible = false;
                lblTutOrani.Visible = false;
                btnKriterAgirlikKaydet.Visible = false;
                btnAhpAyrintiCozGoster.Visible = false;
                lblUyari.Visible = false;
                panel52.Visible = false;
                panel81.Visible = false;
                listBoxKriter.Items.Clear();
                listBoxAlternatif.Items.Clear();
                dataGridViewAgirlik.Columns.Clear();
                dataGridViewKararMat.Columns.Clear();
                dataGridViewNormalize.Columns.Clear();
                dataGridViewBaskinlikSkoru.Columns.Clear();
                dataGridViewGenelSkor.Columns.Clear();
                dataGridViewKarsilastirmaMat.Columns.Clear();
                dataGridViewKriterAgirliklari.Columns.Clear();
                lblTutOrani.Text = "";
                dataGridViewAyrintiKmat.Columns.Clear();
                dataGridViewC.Columns.Clear();
                dataGridViewWVektörü.Columns.Clear();
                dataGridViewDVektör.Columns.Clear();
                lblCI.Text = "";
                lblRI.Text = "";
                lblTutOrani.Text = "";
                dataGridViewSonucKararMat.Columns.Clear();
                dataGridViewSonucNormalizeMat.Columns.Clear();
                dataGridViewSonucKriterAgirlik.Columns.Clear();
                dataGridViewSonucGoreliAgirlik.Columns.Clear();
                dataGridViewSonucKarsilastirmaMat.Columns.Clear();
                dataGridViewSonucGenelBaskinlik.Columns.Clear();
                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel3.Controls.Clear();
                kriterler.Clear();
                alternatifler.Clear();
                faydaMaliyet.Clear();
                agirliklar.Clear();
                goreliAgirliklar.Clear();
                maxList.Clear();
                minList.Clear();
                baskinlikSkorlari.Clear();
                tBaskinlikSkorlari.Clear();
                paydaListesi.Clear();
                paydaListesi.Clear();
                kararMatSutunToplam.Clear();
                genelBaskinlikSkorlari.Clear();
                wr = 0;
                wjr = 0;
                wjr1 = 0;
                max = 0;
                min = 0;
                gAgirlikTop = 0;
                wjrToplam = 0;
                rijFark = 0;
                sonuc = 0;
                sonuc1 = 0;
                baskinlikSkoru = 0;
                tbaskinlikSkoru = 0;
                satirSkorToplam = 0;
                satirSkorToplam = 0;
                maxBaskinlikSkoru = 0;
                minBaskinlikSkoru = 0;
                genelBaskinlikSkoru = 0;
                baskinlikSkorum = 0;
                bskor = 0;
                maxGenelBaskinlikSkor = 0;
                gSkor = 0;
                agirlikToplam = 0;
                enİyiAlternatif = "";
                alternatifBskor = "";
                girilenAlternatif = "";
                rbtnDiziboyut = 1;
                rbtnDizi1boyut = 1;
                x = 0;
                y = 0;
                rbtn = 0;
                rbtn1 = 0;
                lamda = 0;
                CI = 0;
                CR = 0;
                //RadioButton[] radioButton;
                // RadioButton[] radioButton1;
            }
            catch (Exception ex)
            {

                MessageBox.Show("Hata!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public void agirlikTemizle()
        {
            try
            {

                pnlYontemSec.Visible = true;
                panel12.Visible = false;

                btnKriterAgirlikKaydet.Visible = false;
                btnAhpAyrintiCozGoster.Visible = false;
                panel82.Visible = false;
                panel83.Visible = false;
                panel80.Visible = false;
                label24.Visible = false;
                lblTutOrani.Visible = false;
                btnKriterAgirlikKaydet.Visible = false;
                btnAhpAyrintiCozGoster.Visible = false;
                lblUyari.Visible = false;
                panel52.Visible = false;
                panel81.Visible = false;

                dataGridViewAgirlik.Columns.Clear();


                dataGridViewBaskinlikSkoru.Columns.Clear();
                dataGridViewGenelSkor.Columns.Clear();
                dataGridViewKarsilastirmaMat.Columns.Clear();
                dataGridViewKriterAgirliklari.Columns.Clear();
                lblTutOrani.Text = "";
                dataGridViewAyrintiKmat.Columns.Clear();
                dataGridViewC.Columns.Clear();
                dataGridViewWVektörü.Columns.Clear();
                dataGridViewDVektör.Columns.Clear();
                lblCI.Text = "";
                lblRI.Text = "";
                lblTutOrani.Text = "";
                dataGridViewSonucKararMat.Columns.Clear();
                dataGridViewSonucNormalizeMat.Columns.Clear();
                dataGridViewSonucKriterAgirlik.Columns.Clear();
                dataGridViewSonucGoreliAgirlik.Columns.Clear();
                dataGridViewSonucKarsilastirmaMat.Columns.Clear();
                dataGridViewSonucGenelBaskinlik.Columns.Clear();
                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel3.Controls.Clear();

                agirliklar.Clear();
                goreliAgirliklar.Clear();


                baskinlikSkorlari.Clear();
                tBaskinlikSkorlari.Clear();
                paydaListesi.Clear();
                paydaListesi.Clear();

                genelBaskinlikSkorlari.Clear();
                wr = 0;
                wjr = 0;
                wjr1 = 0;

                gAgirlikTop = 0;
                wjrToplam = 0;
                rijFark = 0;
                sonuc = 0;
                sonuc1 = 0;
                baskinlikSkoru = 0;
                tbaskinlikSkoru = 0;
                satirSkorToplam = 0;
                satirSkorToplam = 0;
                maxBaskinlikSkoru = 0;
                minBaskinlikSkoru = 0;
                genelBaskinlikSkoru = 0;
                baskinlikSkorum = 0;
                bskor = 0;
                maxGenelBaskinlikSkor = 0;
                gSkor = 0;
                agirlikToplam = 0;
                enİyiAlternatif = "";
                alternatifBskor = "";
                girilenAlternatif = "";
                rbtnDiziboyut = 1;
                rbtnDizi1boyut = 1;
                x = 0;
                y = 0;
                rbtn = 0;
                rbtn1 = 0;
                lamda = 0;
                CI = 0;
                CR = 0;
                //RadioButton[] radioButton;
                // RadioButton[] radioButton1;
            }
            catch (Exception ex)
            {

                MessageBox.Show("Hata!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void dataGridViewKararMat_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int index = dataGridViewKararMat.CurrentCell.ColumnIndex;
                int i = index - 1;
                if (index == 0)
                {
                    MessageBox.Show("Lütfen yönünü değiştirmek istediğiniz kriterin bulunduğu sutundaki değerlerden birinin üzerine tıklayınız!", "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else if (index != 0)
                {
                    if (faydaMaliyet[i].ToString() == rbtnFayda.Text)
                    {
                        faydaMaliyet.RemoveAt(i);
                        faydaMaliyet.Insert(i, rbtnMaliyet.Text);
                        dataGridViewKararMat.Columns[index].HeaderCell.Style.BackColor = Color.Plum;
                        listBoxKriter.Items.RemoveAt(i);
                        listBoxKriter.Items.Insert(i, (kriterler[i].ToString() + "  (" + rbtnMaliyet.Text + ")"));

                    }
                    else if (faydaMaliyet[i].ToString() == rbtnMaliyet.Text)
                    {
                        faydaMaliyet.RemoveAt(i);
                        faydaMaliyet.Insert(i, rbtnFayda.Text);
                        dataGridViewKararMat.Columns[index].HeaderCell.Style.BackColor = Color.LightBlue;
                        listBoxKriter.Items.RemoveAt(i);
                        listBoxKriter.Items.Insert(i, (kriterler[i].ToString() + "  (" + rbtnFayda.Text + ")"));

                    }
                }
                else
                {

                    MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception)
            {
            }

        }
        private void btnAlternatifDuzenle_Click(object sender, EventArgs e)
        {
            duzenleIndex = listBoxAlternatif.SelectedIndex;
            txtAlternatif.Text = alternatifler[duzenleIndex].ToString();
            btnAlternatifEkle.Text = "Güncelle";
            btnAlternatifEkle.Font = new Font("Bahnschrift Light", 8, FontStyle.Bold);
        }
        public void alternatifDuzenle()
        {
            try
            {
                if (txtAlternatif.Text == "Eklemek istediğiniz alternatifi giriniz")
                {
                    MessageBox.Show("Alternatif girmeden ekleme yapamazsınız !", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else if (txtAlternatif.Text == "")
                {
                    MessageBox.Show("Alternatif girmeden ekleme yapamazsınız !", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    listBoxAlternatif.Items.RemoveAt(duzenleIndex);
                    listBoxAlternatif.Items.Insert(duzenleIndex, txtAlternatif.Text);
                    alternatifler.RemoveAt(duzenleIndex);
                    alternatifler.Insert(duzenleIndex, txtAlternatif.Text);
                    //alternatif eklendikten sonra butonları aktif etsin
                    btnAlternatifSil.Enabled = true;
                    btnAlternatifDuzenle.Enabled = true;
                    //ekledikten sonra textbox ın içini temizleyip imleci oraya fokuslasın

                    txtAlternatif.Clear();
                    txtAlternatif.Focus();
                    pnlAlternatif.Visible = true;
                    pnlKriterAlternatif.Visible = true;
                }

                txtAlternatif.Clear();
                txtAlternatif.Focus();

                btnAlternatifEkle.Text = "Ekle";
                btnAlternatifEkle.Font = new Font("Bahnschrift Light", 9, FontStyle.Bold);
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void txtKriter_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnKriterEkle.PerformClick();
            }
        }
        private void txtAlternatif_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnAlternatifEkle.PerformClick();
            }
        }
        private void txtGirilenAlternatif_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnBaskinlikSkoru.PerformClick();
            }
        }
        private void rbtnFayda_CheckedChanged(object sender, EventArgs e)
        {
            txtKriter.Focus();
        }
        private void rbtnMaliyet_CheckedChanged(object sender, EventArgs e)
        {
            txtKriter.Focus();
        }
        private void btnSimdiOlustur_Click(object sender, EventArgs e)
        {
            tumunuTemizle();
            ileriButonlariniGizle();
            tiklanma = "olustur";
            tabControl1.SelectedTab = tabPageKararMatOlusturma;
        }
        public void kararMatGorunurlukAyarlari()
        {
            pnlKriter.Visible = true;
            pnlAlternatif.Visible = true;
            pnlKriterAlternatif.Visible = true;
            btnKriterDuzenle.Enabled = true;
            btnKriterSil.Enabled = true;
            btnAlternatifDuzenle.Enabled = true;
            btnAlternatifSil.Enabled = true;
        }
        private void btnExcelYukle_Click(object sender, EventArgs e)
        {

            //if (dataGridViewImport.Rows.Count < 149)
            //{
            tumunuTemizle(); //var
            tiklanma = "excel";//var
            ileriButonlariniGizle();
            kararMatGorunurlukAyarlari();//var
            kararMatImport();
            if (dataGridViewImport.Rows.Count != 0)
            {
                importKararMatDoldur();
                kararMatImportListeDoldurma();
                kararMatRenklendir();
                boyutAyarlama();
                tabControl1.SelectedTab = tabPageKararMatrisi;
                gridTasarimSirasiz(dataGridViewKararMat);
                normalizeMatCerceve();
                kararMatrisiToolStripMenuItem1.Visible = true;
            }
            else
            {
                MessageBox.Show("Dosya seçilmedi", "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //}

            //else if (dataGridViewImport.Rows.Count >= 149)
            //{
            //tumunuTemizle(); //var
            //tiklanma = "excel";//var
            //kararMatGorunurlukAyarlari();//var
            //kararMatExcelYukle();
            //if (dataGridViewKararMat.Rows.Count != 0)
            //{
            //    kararMatImportListeDoldurma();
            //    kararMatRenklendir();
            //    boyutAyarlama();
            //    tabControl1.SelectedTab = tabPageKararMatrisi;
            //    normalizeMatCerceve();
            //    kararMatrisiToolStripMenuItem1.Visible = true;
            //}
            //else
            //{
            //    MessageBox.Show("Dosya seçilmedi", "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            //}



        }
        public void kararMatImport()
        {
            try
            {
                var openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
                openFileDialog.FilterIndex = 2;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var workbook = ExcelFile.Load(openFileDialog.FileName);
                    DataGridViewConverter.ExportToDataGridView(workbook.Worksheets.ActiveWorksheet, this.dataGridViewImport, new ExportToDataGridViewOptions() { ColumnHeaders = false });
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Karar matrisi yüklenemedi!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }
        }
        public void sonucKarsilastirmaMatDoldur()
        {
            if (dataGridViewKarsilastirmaMat.Rows.Count > 0)
            {

                pnlSonucKarMat.Visible = true;
                int k = 1;
                dataGridViewSonucKarsilastirmaMat.Rows.Clear();
                dataGridViewSonucKarsilastirmaMat.ColumnCount = dataGridViewKarsilastirmaMat.Columns.Count;


                for (int i = 0; i < kriterler.Count; i++)
                {
                    k = 1;
                    for (int j = 0; j < kriterler.Count; j++)
                    {
                        dataGridViewSonucKarsilastirmaMat.Columns[k].Name = kriterler[j].ToString();
                        k++;
                    }
                }

                for (int j = 0; j < kriterler.Count; j++)
                {
                    dataGridViewSonucKarsilastirmaMat.Rows.Add(kriterler[j].ToString());
                }
                for (int cR = 0; cR < kriterler.Count; cR++)
                {
                    dataGridViewSonucKarsilastirmaMat.Rows[cR].Cells[0].ReadOnly = true;
                }
                for (int i = 0; i < dataGridViewKarsilastirmaMat.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridViewKarsilastirmaMat.Columns.Count; j++)
                    {
                        dataGridViewSonucKarsilastirmaMat.Rows[i].Cells[j].Value = dataGridViewKarsilastirmaMat.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
        }
        public void ahpAyrintiliCozumExcelAktar()
        {
            try
            {
                if (dataGridViewAyrintiKmat.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    xcelApp.Application.Workbooks.Add(Type.Missing);

                    xcelApp.Cells[1, 1] = "KRİTERLER İÇİN İKİLİ KARŞILAŞTIRMA MATRİSİ";
                    for (int i = 1; i < dataGridViewAyrintiKmat.Columns.Count + 1; i++)
                    {
                        xcelApp.Cells[3, i] = dataGridViewAyrintiKmat.Columns[i - 1].HeaderText;
                    }

                    for (int i = 0; i < dataGridViewAyrintiKmat.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewAyrintiKmat.Columns.Count; j++)
                        {
                            xcelApp.Cells[i + 4, j + 1] = dataGridViewAyrintiKmat.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    if (dataGridViewC.Rows.Count > 0)
                    {
                        xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + 7, 1] = "C MATRİSİ";
                        for (int i = 1; i < dataGridViewC.Columns.Count + 1; i++)
                        {

                            xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + 9, i] = dataGridViewC.Columns[i - 1].HeaderText;

                        }
                        int a = 0;
                        for (int i = dataGridViewAyrintiKmat.Rows.Count; i < dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count; i++)
                        {
                            int s = 0;
                            for (int j = 0; j < dataGridViewC.Columns.Count; j++)
                            {
                                xcelApp.Cells[i + 10, j + 1] = dataGridViewC.Rows[a].Cells[s].Value.ToString();
                                s++;
                            }
                            a++;
                        }


                    }

                    if (dataGridViewWVektörü.Rows.Count > 0)
                    {
                        xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + 13, 1] = "AĞIRLIKLAR";
                        for (int i = 1; i < dataGridViewWVektörü.Columns.Count + 1; i++)
                        {

                            xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + 15, i] = dataGridViewWVektörü.Columns[i - 1].HeaderText;

                        }
                        int a = 0;
                        for (int i = dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count; i < dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count; i++)
                        {
                            int s = 0;
                            for (int j = 0; j < dataGridViewWVektörü.Columns.Count; j++)
                            {
                                xcelApp.Cells[i + 16, j + 1] = dataGridViewWVektörü.Rows[a].Cells[s].Value.ToString();
                                s++;
                            }
                            a++;
                        }


                    }


                    if (dataGridViewDVektör.Rows.Count > 0)
                    {
                        xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + 19, 1] = "D VEKTÖRÜ";
                        for (int i = 1; i < dataGridViewDVektör.Columns.Count + 1; i++)
                        {

                            xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + 21, i] = dataGridViewDVektör.Columns[i - 1].HeaderText;

                        }
                        int a = 0;
                        for (int i = dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count; i < dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + dataGridViewDVektör.Rows.Count; i++)
                        {
                            int s = 0;
                            for (int j = 0; j < dataGridViewDVektör.Columns.Count; j++)
                            {
                                xcelApp.Cells[i + 22, j + 1] = dataGridViewDVektör.Rows[a].Cells[s].Value.ToString();
                                s++;
                            }
                            a++;
                        }

                        xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + dataGridViewDVektör.Rows.Count + 26, 1] = "CI: ";
                        xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + dataGridViewDVektör.Rows.Count + 26, 2] = lblCI.Text;

                        xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + dataGridViewDVektör.Rows.Count + 28, 1] = "RI: ";
                        xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + dataGridViewDVektör.Rows.Count + 28, 2] = lblRI.Text;

                        xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + dataGridViewDVektör.Rows.Count + 30, 1] = "TUTARLILIK ORANI: ";
                        xcelApp.Cells[dataGridViewAyrintiKmat.Rows.Count + dataGridViewC.Rows.Count + dataGridViewWVektörü.Rows.Count + dataGridViewDVektör.Rows.Count + 30, 2] = lblTutarlilikOrani.Text;
                    }



                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Excel'e aktarılamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public void sonucExcelAktar()
        {
            try
            {
                if (dataGridViewSonucKararMat.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    xcelApp.Application.Workbooks.Add(Type.Missing);

                    xcelApp.Cells[1, 1] = "KARAR MATRİSİ";
                    for (int i = 1; i < dataGridViewSonucKararMat.Columns.Count + 1; i++)
                    {
                        xcelApp.Cells[2, i] = dataGridViewSonucKararMat.Columns[i - 1].HeaderText;
                    }

                    for (int i = 0; i < dataGridViewSonucKararMat.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridViewSonucKararMat.Columns.Count; j++)
                        {
                            xcelApp.Cells[i + 4, j + 1] = dataGridViewSonucKararMat.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    if (dataGridViewSonucNormalizeMat.Rows.Count > 0)
                    {
                        xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + 8, 1] = "NORMALİZE EDİLMİŞ KARAR MATRİSİ";

                        for (int i = 1; i < dataGridViewSonucNormalizeMat.Columns.Count + 1; i++)
                        {
                            xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + 9, i] = dataGridViewSonucNormalizeMat.Columns[i - 1].HeaderText;

                        }
                        int a = 0;
                        for (int i = dataGridViewSonucKararMat.Rows.Count; i < dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count; i++)
                        {
                            int s = 0;
                            for (int j = 0; j < dataGridViewSonucNormalizeMat.Columns.Count; j++)
                            {
                                xcelApp.Cells[i + 10, j + 1] = dataGridViewSonucNormalizeMat.Rows[a].Cells[s].Value.ToString();
                                s++;
                            }
                            a++;
                        }
                    }

                    if (dataGridViewSonucKriterAgirlik.Rows.Count > 0)
                    {
                        xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + 13, 1] = "KRİTER AĞIRLIKLARI";

                        for (int i = 1; i < dataGridViewSonucKriterAgirlik.Columns.Count + 1; i++)
                        {
                            xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + 14, i] = dataGridViewSonucKriterAgirlik.Columns[i - 1].HeaderText;

                        }
                        int a = 0;
                        for (int i = dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count; i < dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count; i++)
                        {
                            int s = 0;
                            for (int j = 0; j < dataGridViewSonucKriterAgirlik.Columns.Count; j++)
                            {
                                xcelApp.Cells[i + 15, j + 1] = dataGridViewSonucKriterAgirlik.Rows[a].Cells[s].Value.ToString();
                                s++;
                            }
                            a++;
                        }
                    }

                    if (dataGridViewSonucGoreliAgirlik.Rows.Count > 0)
                    {
                        xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + 18, 1] = "GÖRELİ AĞIRLIKLAR";

                        for (int i = 1; i < dataGridViewSonucGoreliAgirlik.Columns.Count + 1; i++)
                        {
                            xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + 19, i] = dataGridViewSonucGoreliAgirlik.Columns[i - 1].HeaderText;

                        }
                        int a = 0;
                        for (int i = dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count; i < dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count; i++)
                        {
                            int s = 0;
                            for (int j = 0; j < dataGridViewSonucGoreliAgirlik.Columns.Count; j++)
                            {
                                xcelApp.Cells[i + 20, j + 1] = dataGridViewSonucGoreliAgirlik.Rows[a].Cells[s].Value.ToString();
                                s++;
                            }
                            a++;
                        }
                    }


                    if (dataGridViewSonucKarsilastirmaMat.Rows.Count > 0)
                    {
                        xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count + 23, 1] = "İKİLİ KARŞILAŞTIRMA MATRİSİ";

                        for (int i = 1; i < dataGridViewSonucKarsilastirmaMat.Columns.Count + 1; i++)
                        {
                            xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count + 24, i] = dataGridViewSonucKarsilastirmaMat.Columns[i - 1].HeaderText;

                        }
                        int a = 0;
                        for (int i = dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count; i < dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count + dataGridViewSonucKarsilastirmaMat.Rows.Count; i++)
                        {
                            int s = 0;
                            for (int j = 0; j < dataGridViewSonucKarsilastirmaMat.Columns.Count; j++)
                            {
                                xcelApp.Cells[i + 25, j + 1] = dataGridViewSonucKarsilastirmaMat.Rows[a].Cells[s].Value.ToString();
                                s++;
                            }
                            a++;
                        }
                    }

                    if (dataGridViewSonucGenelBaskinlik.Rows.Count > 0)
                    {
                        xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count + dataGridViewSonucKarsilastirmaMat.Rows.Count + 28, 1] = "GENEL BASKINLIK SKORLARI";
                        for (int i = 1; i < dataGridViewSonucGenelBaskinlik.Columns.Count + 1; i++)
                        {
                            xcelApp.Cells[dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count + dataGridViewSonucKarsilastirmaMat.Rows.Count + 29, i] = dataGridViewSonucGenelBaskinlik.Columns[i - 1].HeaderText;
                        }
                        int a = 0;
                        for (int i = dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count + dataGridViewSonucKarsilastirmaMat.Rows.Count; i < dataGridViewSonucKararMat.Rows.Count + dataGridViewSonucNormalizeMat.Rows.Count + dataGridViewSonucKriterAgirlik.Rows.Count + dataGridViewSonucGoreliAgirlik.Rows.Count + dataGridViewSonucKarsilastirmaMat.Rows.Count + dataGridViewSonucGenelBaskinlik.Rows.Count; i++)
                        {
                            int s = 0;
                            for (int j = 0; j < dataGridViewSonucGenelBaskinlik.Columns.Count; j++)
                            {
                                xcelApp.Cells[i + 30, j + 1] = dataGridViewSonucGenelBaskinlik.Rows[a].Cells[s].Value.ToString();
                                s++;
                            }
                            a++;
                        }
                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarılamadı!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void btnTumBSkor_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            tumKismiBaskinlikSkorGoruntule();
            tabControl1.SelectedTab = tabPageBaskinlikSkorSonuc;
        }
        public void tumKismiBaskinlikSkorGoruntule()
        {
            try
            {
                int gridSayi = alternatifler.Count;

                for (int i = 0; i < alternatifler.Count; i++)
                {
                    Button btnA = new Button();
                    btnA.Name = "btnA" + i.ToString();
                    btnA.AutoSize = false;
                    btnA.Text = alternatifler[i].ToString();
                    btnA.Width = 150;
                    btnA.Height = 35;
                    btnA.BackColor = Color.Thistle;
                    btnA.ForeColor = Color.White;
                    btnA.FlatStyle = FlatStyle.Popup;
                    btnA.Font = new Font(" Bahnschrift Light", 8, FontStyle.Bold);
                    btnA.Click += new EventHandler(btnA_Click);
                    flowLayoutPanel1.Controls.Add(btnA);
                }
                Button btnH = new Button();
                btnH.Name = "btnTümü";
                btnH.AutoSize = false;
                btnH.Text = "HEPSİNİ GÖSTER";
                btnH.Width = 150;
                btnH.Height = 35;
                btnH.Font = new Font(" Bahnschrift Light", 8, FontStyle.Bold);
                btnH.BackColor = Color.Thistle;
                btnH.ForeColor = Color.White;
                btnH.FlatStyle = FlatStyle.Popup;
                btnH.Click += new EventHandler(btnH_Click);
                flowLayoutPanel1.Controls.Add(btnH);
            }
            catch (Exception)
            {

                MessageBox.Show("Kısmi baskınlık skorları görüntülenemedi!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


        }

        //public event System.Windows.Forms.MouseEventHandler MouseMove;
        //public event System.Windows.Forms.MouseEventHandler MouseUp;
        //public event System.Windows.Forms.MouseEventHandler MouseHover;
        public void ahpTasarim()
        {
            //faktöriyel hesabı
            int f = 1;
            for (int i = 1; i <= kriterler.Count - 1; i++)
            {
                f = i * f;
            }

            pnlAhpTasarim.Controls.Clear();
            panelSayiAhp.Controls.Clear();
            for (int i = 0; i < 9; i++)
            {
                Label lbl = new Label();
                int yi = 9 - i;
                lbl.Name = "lbls" + yi.ToString();
                lbl.Text = yi.ToString();
                lbl.Top = 0;
                lbl.Left = 285 + (22 * i);
                lbl.Width = 12;
                lbl.Height = 15;
                lbl.Font = new Font("Palatino Linotype", 7, FontStyle.Bold);
                panelSayiAhp.Controls.Add(lbl);

            }

            for (int i = 0; i < 8; i++)
            {
                Label lbl = new Label();
                int yi = i + 2;
                lbl.Name = "lblk" + yi.ToString();
                lbl.Text = yi.ToString();
                lbl.Top = 0;
                lbl.Left = 485 + (22 * i);
                lbl.Width = 12;
                lbl.Height = 15;
                lbl.Font = new Font("Palatino Linotype", 7, FontStyle.Bold);
                panelSayiAhp.Controls.Add(lbl);

            }

            ///BOYUT///////////////////////////////////

            for (int i = 1; i < kriterler.Count; i++)
            {
                for (int j = i; j < kriterler.Count; j++)
                {
                    for (int rb = 1; rb < 10; rb++) // 9 tane 
                    {
                        rbtnDiziboyut++;
                    }
                }

            }

            for (int i = 1; i < kriterler.Count; i++)
            {
                for (int j = i; j < kriterler.Count; j++)
                {
                    for (int rb = 1; rb < 9; rb++) //8 tane
                    {
                        rbtnDizi1boyut++;
                    }
                }

            }

            radioButton = new RadioButton[rbtnDiziboyut];
            radioButton1 = new RadioButton[rbtnDizi1boyut];
            ///////////////////////////////////
            int k = 0;
            int enBuyuk = 0;
            int enBuyuk2 = 0;
            for (int i = 1; i < kriterler.Count; i++)
            {
                for (int j = i; j < kriterler.Count; j++)
                {
                    Label lbl = new Label();
                    lbl.Name = "lblRow" + i + j;
                    lbl.Text = kriterler[i - 1].ToString();
                    //lbl.Top = 53 + (29 * k);
                    lbl.Top = 8 + (29 * k);

                    if (lbl.Text.Length > enBuyuk)
                    {
                        enBuyuk = lbl.Text.Length;
                    }

                    if (enBuyuk <= 2)
                    {
                        lbl.Left = 243;
                        lbl.Width = 33;
                    }
                    else if (enBuyuk >= 2 && enBuyuk <= 5)
                    {
                        lbl.Left = 220;
                        lbl.Width = 53;
                    }
                    else
                    {
                        lbl.Left = 218 - (enBuyuk * 6);
                        //lbl.Width = 45 * (enBuyuk / 6);
                        lbl.Width = enBuyuk * 8;

                    }


                    lbl.Height = 17;
                    lbl.Font = new Font("Palatino Linotype", 9, FontStyle.Bold);
                    pnlAhpTasarim.Controls.Add(lbl);

                    Label lbl1 = new Label();
                    lbl1.Name = "lblCol" + i + j;
                    lbl1.Text = kriterler[j].ToString();
                    //lbl1.Top = 53 + (29 * k);
                    lbl1.Top = 8 + (29 * k);

                    if (lbl1.Text.Length > enBuyuk2)
                    {
                        enBuyuk2 = lbl1.Text.Length;
                    }

                    if (enBuyuk2 <= 5)
                    {
                        lbl1.Left = 666;
                        lbl1.Width = 53;
                    }
                    else
                    {
                        lbl1.Left = 666;
                        //lbl1.Left = 666 + (enBuyuk2 * 7);
                        //lbl1.Width = 53 * (enBuyuk2 / 6);
                        lbl1.Width = enBuyuk2 * 8;
                    }

                    lbl1.Height = 17;
                    lbl1.Font = new Font("Palatino Linotype", 9, FontStyle.Bold);
                    pnlAhpTasarim.Controls.Add(lbl1);

                    GroupBox groupBox = new GroupBox();
                    int J = j + 1;
                    groupBox.Name = "groupBox" + i + J;
                    groupBox.Text = "";
                    //groupBox.Top = 45 + (29 * k);
                    groupBox.Top = (29 * k);
                    groupBox.Left = 277;
                    groupBox.Width = 380;
                    groupBox.Height = 30;

                    int no = 9;
                    for (int rb = 1; rb < 10; rb++)
                    {
                        int s = rb - 1;
                        radioButton[rbtn] = new RadioButton();
                        //radioButton[rbtn].Name = "rBtnR" + i + J + rb;
                        radioButton[rbtn].Name = no.ToString();
                        radioButton[rbtn].Text = i.ToString();
                        radioButton[rbtn].ForeColor = Color.White;
                        radioButton[rbtn].Top = 10;
                        radioButton[rbtn].Left = 7 + (22 * s);
                        radioButton[rbtn].Width = 14;
                        radioButton[rbtn].Height = 13;
                        radioButton[rbtn].CheckedChanged += new EventHandler(RadioChange);
                        radioButton[rbtn].MouseHover += new EventHandler(radioButton_MouseHover);
                        radioButton[rbtn].MouseLeave += new EventHandler(radioButton_MouseLeave);
                        groupBox.Controls.Add(radioButton[rbtn]);
                        no--;
                        rbtn++;
                    }
                    int no2 = 2;
                    for (int rb = 1; rb < 9; rb++)
                    {
                        int s = rb - 1;
                        radioButton1[rbtn1] = new RadioButton();
                        //radioButton1[rbtn1].Name = "rBtnC" + i + J + rb.ToString();
                        radioButton1[rbtn1].Name = no2.ToString();
                        radioButton1[rbtn1].Text = "";
                        radioButton1[rbtn1].ForeColor = Color.White;
                        radioButton1[rbtn1].Top = 10;
                        radioButton1[rbtn1].Left = 206 + (22 * s);
                        radioButton1[rbtn1].Width = 14;
                        radioButton1[rbtn1].Height = 13;
                        radioButton1[rbtn1].CheckedChanged += new EventHandler(RadioChange);
                        radioButton1[rbtn1].MouseHover += new EventHandler(radioButton1_MouseHover);
                        radioButton1[rbtn1].MouseLeave += new EventHandler(radioButton1_MouseLeave);
                        groupBox.Controls.Add(radioButton1[rbtn1]);
                        no2++;
                        rbtn1++;
                    }
                    pnlAhpTasarim.Controls.Add(groupBox);
                    k++;
                }
            }

            Button btnKarsilastirmaMatOlustur = new Button();
            btnKarsilastirmaMatOlustur.Name = "btnKararMatOlustur";
            btnKarsilastirmaMatOlustur.AutoSize = false;
            btnKarsilastirmaMatOlustur.Text = "KAYDET";
            btnKarsilastirmaMatOlustur.Top = 7;
            btnKarsilastirmaMatOlustur.Left = 370;
            btnKarsilastirmaMatOlustur.Width = 200;
            btnKarsilastirmaMatOlustur.Height = 36;
            btnKarsilastirmaMatOlustur.Font = new Font(" Bahnschrift Light", 9, FontStyle.Bold);
            btnKarsilastirmaMatOlustur.BackColor = Color.Gray;
            btnKarsilastirmaMatOlustur.ForeColor = Color.White;
            btnKarsilastirmaMatOlustur.FlatStyle = FlatStyle.Popup;
            btnKarsilastirmaMatOlustur.Click += new EventHandler(btnKarsilastirmaMatOlustur_Click);
            panel62.Controls.Add(btnKarsilastirmaMatOlustur);
        }
        private void radioButton_MouseHover(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            //radioButton.BackColor = Color.PaleGreen;
            bilgiMesajiRadioButton(radioButton.Name, radioButton);
        }
        private void radioButton_MouseLeave(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            //radioButton.BackColor = Color.White;
            bilgiMesajiRadioButton(radioButton.Name, radioButton);

        }
        private void radioButton1_MouseHover(object sender, EventArgs e)
        {
            RadioButton radioButton1 = (RadioButton)sender;
            //radioButton1.BackColor = Color.PaleGreen;
            bilgiMesajiRadioButton(radioButton1.Name, radioButton1);

        }
        private void radioButton1_MouseLeave(object sender, EventArgs e)
        {
            RadioButton radioButton1 = (RadioButton)sender;
            //radioButton1.BackColor = Color.White;
            bilgiMesajiRadioButton(radioButton1.Name, radioButton1);

        }
        private void btnA_Click(object sender, EventArgs e)
        {
            try
            {
                baskinlikSkorButon = 1;
                Button btnA = (Button)sender;
                btnA.BackColor = Color.LightSteelBlue;
                flowLayoutPanel3.Controls.Clear();
                for (int i = 0; i < alternatifler.Count; i++)
                {
                    string a = alternatifler[i].ToString();
                    if (btnA.Text == a)
                    {
                        /*DataGridView*/
                        dgvTumBSkor = new DataGridView();
                        dgvTumBSkor.Name = "dgvA" + i.ToString();
                        dgvTumBSkor.Width = 590;
                        dgvTumBSkor.Height = 180;
                        dgvTumBSkor.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                        dgvTumBSkor.AllowUserToAddRows = false;
                        dgvTumBSkor.AllowUserToDeleteRows = false;
                        dgvTumBSkor.BackgroundColor = Color.White;
                        gridTasarim(dgvTumBSkor);
                        /////////////////////////////////////////////////////////////////////////////////////////////////
                        dgvTumBSkor.ColumnCount = kriterler.Count + 2;
                        dgvTumBSkor.Columns[0].Name = " ";
                        dgvTumBSkor.Columns[kriterler.Count + 1].Name = "Toplam";
                        int j = 1;
                        for (int ii = 0; ii < kriterler.Count; ii++)
                        {
                            dgvTumBSkor.Columns[j].Name = "P" + (ii + 1).ToString() + "(Ai,Ai')";
                            j++;
                        }

                        for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                        {
                            if (iUssu == i)
                            {
                                continue;
                            }

                            dgvTumBSkor.Rows.Add(("(" + alternatifler[i].ToString()) + "," + alternatifler[iUssu].ToString() + ")");

                        }


                        //label7.Text = alternatifBskor + " ALTERNATİFİNİN " + alternatifBskor + "' ALTERNATİFLERİNE BASKINLIK SKORU:";
                        for (int k = 1; k < kriterler.Count + 1; k++)
                        {
                            for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                            {
                                if (iUssu == i)
                                {
                                    continue;
                                }

                                rijFark = Convert.ToDouble(dataGridViewNormalize.Rows[i].Cells[k].Value) - Convert.ToDouble(dataGridViewNormalize.Rows[iUssu].Cells[k].Value);
                                wjr1 = Convert.ToDouble(goreliAgirliklar[k - 1]);
                                int indis = kriterler.Count;
                                wjrToplam = Convert.ToDouble(goreliAgirliklar[indis]);

                                if (rijFark > 0)
                                {
                                    if (wjrToplam == 0)
                                    {
                                        sonuc = (wjr1 * rijFark);
                                        sonuc = Math.Sqrt(sonuc);

                                        if (i > iUssu)
                                        {
                                            dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc;
                                        }
                                        else if (i < iUssu)
                                        {
                                            dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                        }
                                    }

                                    else
                                    {
                                        sonuc = (wjr1 * rijFark) / wjrToplam;
                                        sonuc = Math.Sqrt(sonuc);
                                        if (i > iUssu)
                                        {
                                            dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc;
                                        }
                                        else if (i < iUssu)
                                        {
                                            dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                        }
                                    }

                                }

                                else if (rijFark == 0)
                                {

                                    sonuc = 0;
                                    if (i > iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc;
                                    }
                                    else if (i < iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                    }

                                }

                                else //rijFark <0 durumu
                                {

                                    if (wjr1 == 0)
                                    {
                                        if (θ == 0)
                                        {
                                            sonuc = 0;

                                            if (i > iUssu)
                                            {
                                                dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc;
                                            }
                                            else if (i < iUssu)
                                            {
                                                dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                            }
                                        }
                                        else
                                        {
                                            sonuc = (wjrToplam * rijFark) * -1;
                                            sonuc = (Math.Sqrt(sonuc)) * -Convert.ToDouble(1 / θ);


                                            if (i > iUssu)
                                            {
                                                dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc;
                                            }
                                            else if (i < iUssu)
                                            {
                                                dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                            }
                                        }
                                    }

                                    else
                                    {
                                        if (θ == 0)
                                        {
                                            sonuc = 0;

                                            if (i > iUssu)
                                            {
                                                dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc;
                                            }
                                            else if (i < iUssu)
                                            {
                                                dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                            }
                                        }
                                        else
                                        {
                                            sonuc = ((wjrToplam * rijFark) / wjr1) * -1;
                                            sonuc = (Math.Sqrt(sonuc)) * -Convert.ToDouble(1 / θ);

                                            if (i > iUssu)
                                            {
                                                dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc;
                                            }
                                            else if (i < iUssu)
                                            {
                                                dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc;
                                            }
                                        }
                                    }

                                }
                            }

                        }

                        for (int s = 0; s < alternatifler.Count - 1; s++)
                        {
                            for (int c = 1; c < kriterler.Count + 1; c++)
                            {
                                satirSkorToplam += Convert.ToDouble(dgvTumBSkor.Rows[s].Cells[c].Value);
                            }
                            dgvTumBSkor.Rows[s].Cells[kriterler.Count + 1].Value = satirSkorToplam;
                            baskinlikSkoru += satirSkorToplam;
                            satirSkorToplam = 0;
                        }

                        dgvTumBSkor.Rows.Add("S(Ai,Ai')");
                        dgvTumBSkor.Rows[alternatifler.Count - 1].Cells[kriterler.Count + 1].Value = baskinlikSkoru;

                        baskinlikSkoru = 0;
                        flowLayoutPanel3.Controls.Add(dgvTumBSkor);
                    }

                    else
                    {
                        continue;
                    }
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void btnH_Click(object sender, EventArgs e)
        {
            try
            {
                baskinlikSkorButon = 2;
                // tumunu excel aktarma için
                int alternatifSay = alternatifler.Count;
                dgvTumBSkorDizi = new DataGridView[alternatifSay];
                //

                Button btnH = (Button)sender;
                btnH.BackColor = Color.LightSteelBlue;
                flowLayoutPanel3.Controls.Clear();
                for (int i = 0; i < alternatifler.Count; i++)
                {
                    // excel aktarma için
                    dgvTumBSkorDizi[i] = new DataGridView();
                    dgvTumBSkorDizi[i].Name = "dgvAB" + i.ToString();
                    dgvTumBSkorDizi[i].Width = 590;
                    dgvTumBSkorDizi[i].Height = 25 * alternatifler.Count;
                    dgvTumBSkorDizi[i].ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    dgvTumBSkorDizi[i].AllowUserToAddRows = false;
                    dgvTumBSkorDizi[i].AllowUserToDeleteRows = false;
                    dgvTumBSkorDizi[i].BackgroundColor = Color.White;
                    gridTasarim(dgvTumBSkorDizi[i]);
                    //bu kodlar işe yaramadı :( tasarımı değişmiyor :( :( :( :( :( :( :(
                    dgvTumBSkorDizi[i].RowHeadersVisible = false;  //ilk sutunu gizleme
                    dgvTumBSkorDizi[i].BorderStyle = BorderStyle.None;
                    dgvTumBSkorDizi[i].AlternatingRowsDefaultCellStyle.BackColor = Color.LightSlateGray;//varsayılan arka plan rengi verme
                    dgvTumBSkorDizi[i].DefaultCellStyle.SelectionBackColor = Color.Silver; //seçilen hücrenin arkaplan rengini belirleme
                    dgvTumBSkorDizi[i].DefaultCellStyle.SelectionForeColor = Color.White;  //seçilen hücrenin yazı rengini belirleme
                    dgvTumBSkorDizi[i].EnableHeadersVisualStyles = false; //başlık özelliğini değiştirme
                    dgvTumBSkorDizi[i].ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; //başlık çizgilerini ayarlama
                    dgvTumBSkorDizi[i].ColumnHeadersDefaultCellStyle.BackColor = Color.Gray; //başlık arkaplan rengini belirleme
                    dgvTumBSkorDizi[i].ColumnHeadersDefaultCellStyle.ForeColor = Color.White;  //başlık yazı rengini belirleme
                    dgvTumBSkorDizi[i].AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;                                                                   // datagridview.SelectionMode = DataGridViewSelectionMode.FullRowSelect; //satırı tamamen seçmeyi sağlama
                    dgvTumBSkorDizi[i].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;  //her hangi bir sutunun genişliğini o sutunda yer alan en  uzun yazının genişliğine göre ayarlama
                    dgvTumBSkorDizi[i].AllowUserToAddRows = false;  //ilk sutunu gizleme
                    dgvTumBSkorDizi[i].AllowUserToOrderColumns = false;

                    /////////////////////////////////////////////////////////////////////////////////////////////////

                    dgvTumBSkorDizi[i].ColumnCount = kriterler.Count + 2;
                    dgvTumBSkorDizi[i].Columns[0].Name = " ";
                    dgvTumBSkorDizi[i].Columns[kriterler.Count + 1].Name = "Toplam";
                    int col = 1;
                    for (int ii = 0; ii < kriterler.Count; ii++)
                    {
                        dgvTumBSkorDizi[i].Columns[col].Name = "P" + (ii + 1).ToString() + "(Ai,Ai')";
                        col++;
                    }

                    for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                    {
                        if (iUssu == i)
                        {
                            continue;
                        }

                        dgvTumBSkorDizi[i].Rows.Add(("(" + alternatifler[i].ToString()) + "," + alternatifler[iUssu].ToString() + ")");

                    }


                    //label7.Text = alternatifBskor + " ALTERNATİFİNİN " + alternatifBskor + "' ALTERNATİFLERİNE BASKINLIK SKORU:";
                    for (int k = 1; k < kriterler.Count + 1; k++)
                    {
                        for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                        {
                            if (iUssu == i)
                            {
                                continue;
                            }

                            rijFark = Convert.ToDouble(dataGridViewNormalize.Rows[i].Cells[k].Value) - Convert.ToDouble(dataGridViewNormalize.Rows[iUssu].Cells[k].Value);
                            wjr1 = Convert.ToDouble(goreliAgirliklar[k - 1]);
                            int indis = kriterler.Count;
                            wjrToplam = Convert.ToDouble(goreliAgirliklar[indis]);

                            if (rijFark > 0)
                            {
                                if (wjrToplam == 0)
                                {
                                    sonuc = (wjr1 * rijFark);
                                    sonuc = Math.Sqrt(sonuc);

                                    if (i > iUssu)
                                    {
                                        dgvTumBSkorDizi[i].Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                    }
                                    else if (i < iUssu)
                                    {
                                        dgvTumBSkorDizi[i].Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                    }
                                }

                                else
                                {
                                    sonuc = (wjr1 * rijFark) / wjrToplam;
                                    sonuc = Math.Sqrt(sonuc);
                                    if (i > iUssu)
                                    {
                                        dgvTumBSkorDizi[i].Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                    }
                                    else if (i < iUssu)
                                    {
                                        dgvTumBSkorDizi[i].Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                    }
                                }

                            }

                            else if (rijFark == 0)
                            {

                                sonuc = 0;
                                if (i > iUssu)
                                {
                                    dgvTumBSkorDizi[i].Rows[iUssu].Cells[k].Value = sonuc;
                                }
                                else if (i < iUssu)
                                {
                                    dgvTumBSkorDizi[i].Rows[iUssu - 1].Cells[k].Value = sonuc;
                                }

                            }

                            else/* if (rijFark <0) /* durumu*/
                            {

                                if (wjr1 == 0)
                                {
                                    if (θ == 0)
                                    {
                                        sonuc = 0;

                                        if (i > iUssu)
                                        {
                                            dgvTumBSkorDizi[i].Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                        }
                                        else if (i < iUssu)
                                        {
                                            dgvTumBSkorDizi[i].Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                        }
                                    }

                                    else
                                    {
                                        sonuc = (wjrToplam * rijFark) * -1;
                                        sonuc = (Math.Sqrt(sonuc)) * -Convert.ToDouble(1 / θ);


                                        if (i > iUssu)
                                        {
                                            dgvTumBSkorDizi[i].Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                        }
                                        else if (i < iUssu)
                                        {
                                            dgvTumBSkorDizi[i].Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                        }
                                    }

                                }

                                else
                                {
                                    if (θ == 0)
                                    {
                                        sonuc = 0;


                                        if (i > iUssu)
                                        {
                                            dgvTumBSkorDizi[i].Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                        }
                                        else if (i < iUssu)
                                        {
                                            dgvTumBSkorDizi[i].Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                        }
                                    }

                                    else
                                    {

                                        sonuc = ((wjrToplam * rijFark) / wjr1) * -1;
                                        sonuc = (Math.Sqrt(sonuc)) * -Convert.ToDouble(1 / θ);

                                        if (i > iUssu)
                                        {
                                            dgvTumBSkorDizi[i].Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                        }
                                        else if (i < iUssu)
                                        {
                                            dgvTumBSkorDizi[i].Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                        }
                                    }

                                }



                            }

                        }

                    }

                    for (int s = 0; s < alternatifler.Count - 1; s++)
                    {
                        for (int c = 1; c < kriterler.Count + 1; c++)
                        {
                            satirSkorToplam += Convert.ToDouble(dgvTumBSkorDizi[i].Rows[s].Cells[c].Value);
                        }
                        dgvTumBSkorDizi[i].Rows[s].Cells[kriterler.Count + 1].Value = satirSkorToplam.ToString();
                        baskinlikSkoru += satirSkorToplam;
                        satirSkorToplam = 0;
                    }

                    dgvTumBSkorDizi[i].Rows.Add("S(Ai,Ai')");
                    dgvTumBSkorDizi[i].Rows[alternatifler.Count - 1].Cells[kriterler.Count + 1].Value = baskinlikSkoru.ToString();

                    baskinlikSkoru = 0;



                    //excel aktarma burada son ///////////////////////////////////////////
                    //////////////////////////////////////////////////////////////////////

                    /*DataGridView*/
                    dgvTumBSkor = new DataGridView();
                    dgvTumBSkor.Name = "dgvA" + i.ToString();
                    dgvTumBSkor.Width = 590;
                    dgvTumBSkor.Height = 25 * alternatifler.Count;
                    dgvTumBSkor.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    dgvTumBSkor.AllowUserToAddRows = false;
                    dgvTumBSkor.AllowUserToDeleteRows = false;
                    dgvTumBSkor.BackgroundColor = Color.White;
                    gridTasarim(dgvTumBSkor);
                    //bu kodlar işe yaramadı :( tasarımı değişmiyor :( :( :( :( :( :( :(
                    dgvTumBSkor.RowHeadersVisible = false;  //ilk sutunu gizleme
                    dgvTumBSkor.BorderStyle = BorderStyle.None;
                    dgvTumBSkor.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSlateGray;//varsayılan arka plan rengi verme
                    dgvTumBSkor.DefaultCellStyle.SelectionBackColor = Color.Silver; //seçilen hücrenin arkaplan rengini belirleme
                    dgvTumBSkor.DefaultCellStyle.SelectionForeColor = Color.White;  //seçilen hücrenin yazı rengini belirleme
                    dgvTumBSkor.EnableHeadersVisualStyles = false; //başlık özelliğini değiştirme
                    dgvTumBSkor.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None; //başlık çizgilerini ayarlama
                    dgvTumBSkor.ColumnHeadersDefaultCellStyle.BackColor = Color.Gray; //başlık arkaplan rengini belirleme
                    dgvTumBSkor.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;  //başlık yazı rengini belirleme
                    dgvTumBSkor.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;                                                                   // datagridview.SelectionMode = DataGridViewSelectionMode.FullRowSelect; //satırı tamamen seçmeyi sağlama
                    dgvTumBSkor.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;  //her hangi bir sutunun genişliğini o sutunda yer alan en  uzun yazının genişliğine göre ayarlama
                    dgvTumBSkor.AllowUserToAddRows = false;  //ilk sutunu gizleme
                    dgvTumBSkor.AllowUserToOrderColumns = false;
                    dgvTumBSkor.ColumnHeadersHeight = 30;
                    /////////////////////////////////////////////////////////////////////////////////////////////////

                    dgvTumBSkor.ColumnCount = kriterler.Count + 2;
                    dgvTumBSkor.Columns[0].Name = " ";
                    dgvTumBSkor.Columns[kriterler.Count + 1].Name = "Toplam";
                    int j = 1;
                    for (int ii = 0; ii < kriterler.Count; ii++)
                    {
                        dgvTumBSkor.Columns[j].Name = "P" + (ii + 1).ToString() + "(Ai,Ai')";
                        j++;
                    }

                    for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                    {
                        if (iUssu == i)
                        {
                            continue;
                        }

                        dgvTumBSkor.Rows.Add(("(" + alternatifler[i].ToString()) + "," + alternatifler[iUssu].ToString() + ")");

                    }


                    //label7.Text = alternatifBskor + " ALTERNATİFİNİN " + alternatifBskor + "' ALTERNATİFLERİNE BASKINLIK SKORU:";
                    for (int k = 1; k < kriterler.Count + 1; k++)
                    {
                        for (int iUssu = 0; iUssu < alternatifler.Count; iUssu++)
                        {
                            if (iUssu == i)
                            {
                                continue;
                            }

                            rijFark = Convert.ToDouble(dataGridViewNormalize.Rows[i].Cells[k].Value) - Convert.ToDouble(dataGridViewNormalize.Rows[iUssu].Cells[k].Value);
                            wjr1 = Convert.ToDouble(goreliAgirliklar[k - 1]);
                            int indis = kriterler.Count;
                            wjrToplam = Convert.ToDouble(goreliAgirliklar[indis]);

                            if (rijFark > 0)
                            {
                                if (wjrToplam == 0)
                                {
                                    sonuc = (wjr1 * rijFark);
                                    sonuc = Math.Sqrt(sonuc);

                                    if (i > iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                    }
                                    else if (i < iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                    }
                                }

                                else
                                {
                                    sonuc = (wjr1 * rijFark) / wjrToplam;
                                    sonuc = Math.Sqrt(sonuc);
                                    if (i > iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                    }
                                    else if (i < iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                    }
                                }

                            }

                            else if (rijFark == 0)
                            {

                                sonuc = 0;
                                if (i > iUssu)
                                {
                                    dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                }
                                else if (i < iUssu)
                                {
                                    dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                }

                            }

                            else //rijFark <0 durumu
                            {

                                if (wjr1 == 0)
                                {
                                    sonuc = (wjrToplam * rijFark) * -1;
                                    sonuc = (Math.Sqrt(sonuc)) * -Convert.ToDouble(1 / θ);


                                    if (i > iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                    }
                                    else if (i < iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                    }
                                }

                                else
                                {

                                    sonuc = ((wjrToplam * rijFark) / wjr1) * -1;
                                    sonuc = (Math.Sqrt(sonuc)) * -Convert.ToDouble(1 / θ);

                                    if (i > iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu].Cells[k].Value = sonuc.ToString();
                                    }
                                    else if (i < iUssu)
                                    {
                                        dgvTumBSkor.Rows[iUssu - 1].Cells[k].Value = sonuc.ToString();
                                    }

                                }

                            }
                        }

                    }

                    for (int s = 0; s < alternatifler.Count - 1; s++)
                    {
                        for (int c = 1; c < kriterler.Count + 1; c++)
                        {
                            satirSkorToplam += Convert.ToDouble(dgvTumBSkor.Rows[s].Cells[c].Value);
                        }
                        dgvTumBSkor.Rows[s].Cells[kriterler.Count + 1].Value = satirSkorToplam.ToString();
                        baskinlikSkoru += satirSkorToplam;
                        satirSkorToplam = 0;
                    }

                    dgvTumBSkor.Rows.Add("S(Ai,Ai')");
                    dgvTumBSkor.Rows[alternatifler.Count - 1].Cells[kriterler.Count + 1].Value = baskinlikSkoru.ToString();

                    baskinlikSkoru = 0;

                    flowLayoutPanel3.Controls.Add(dgvTumBSkor);

                    Panel p = new Panel();
                    p.Name = "p" + i.ToString();
                    p.Width = 25;
                    p.Height = 25 * alternatifler.Count;
                    flowLayoutPanel3.Controls.Add(p);

                }
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

        }
        private void btnTumKismiBskorExcelAktar_Click(object sender, EventArgs e)
        {
            tumBskorExcelAktar();
        }
        public void tumBskorExcelAktar()
        {
            if (baskinlikSkorButon == 2)
            {
                try
                {
                    ExcelPackage package = new ExcelPackage();
                    package.Workbook.Worksheets.Add("Worksheets1");
                    OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    worksheet.Cells[1, 1].Value = "KISMİ BASKINLIK SKORLARI";
                    worksheet.Cells[1, 1, 1, 4].Merge = true;
                    int satir = 2;
                    for (int f = 0; f < alternatifler.Count; f++)
                    {
                        satir++;
                        var columns = dgvTumBSkorDizi[f].Columns;
                        for (int i = 1; i < columns.Count; i++)
                        {
                            worksheet.Cells[satir, i + 1].Value = "P" + i + "(Ai,Ai')";
                        }
                        worksheet.Cells[satir, columns.Count].Value = "Toplam";
                        satir++;

                        var rows = dgvTumBSkorDizi[f].Rows;
                        for (int i = 0; i < rows.Count; i++)
                        {
                            if (rows[i].Cells[0] != null)
                            {
                                for (int j = 0; j < rows[i].Cells.Count; j++)
                                {
                                    worksheet.Cells[satir, j + 1].Value = rows[i].Cells[j].Value;
                                }
                                satir++;
                            }
                        }

                    }
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                    saveFileDialog.ShowDialog();

                    Stream stream = saveFileDialog.OpenFile();
                    package.SaveAs(stream);
                    stream.Close();

                    MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
                }
                catch (Exception)
                {
                    MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            else if (baskinlikSkorButon == 1)
            {
                try
                {
                    ExcelPackage package = new ExcelPackage();
                    package.Workbook.Worksheets.Add("Worksheets1");
                    OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    worksheet.Cells[1, 1].Value = "KISMİ BASKINLIK SKORU";
                    worksheet.Cells[1, 1, 1, 4].Merge = true;
                    int satir = 2;
                    satir++;
                    var columns = dgvTumBSkor.Columns;
                    for (int i = 1; i < columns.Count; i++)
                    {
                        worksheet.Cells[satir, i + 1].Value = "P" + i + "(Ai,Ai')";
                    }
                    worksheet.Cells[satir, columns.Count].Value = "Toplam";
                    satir++;

                    var rows = dgvTumBSkor.Rows;
                    for (int i = 0; i < rows.Count; i++)
                    {
                        if (rows[i].Cells[0] != null)
                        {
                            for (int j = 0; j < rows[i].Cells.Count; j++)
                            {
                                worksheet.Cells[satir, j + 1].Value = rows[i].Cells[j].Value;
                            }
                            satir++;
                        }
                    }


                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                    saveFileDialog.ShowDialog();

                    Stream stream = saveFileDialog.OpenFile();
                    package.SaveAs(stream);
                    stream.Close();

                    MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
                }
                catch (Exception)
                {
                    MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            //if (baskinlikSkorButon == 2)
            //{
            //    try
            //    {
            //        var saveFileDialog = new SaveFileDialog();
            //        saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            //        saveFileDialog.FilterIndex = 3;

            //        if (saveFileDialog.ShowDialog() == DialogResult.OK)
            //        {

            //            var workbook = new ExcelFile();
            //            var worksheet = workbook.Worksheets.Add("Sheet1");
            //            int say = 0;
            //            for (int i = 0; i < alternatifler.Count; i++)
            //            {

            //                DataGridViewConverter.ImportFromDataGridView(worksheet, this.dgvTumBSkorDizi[i], new ImportFromDataGridViewOptions() { ColumnHeaders = true, StartRow = say });
            //                say += dgvTumBSkorDizi[i].Rows.Count + 4;

            //            }
            //            workbook.Save(saveFileDialog.FileName);
            //        }
            //    }
            //    catch
            //    {
            //        MessageBox.Show("Excel'e aktarma işlemi başarısız! En fazla 150 satır aktarılabilmektedir.", "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        return;
            //    }
            //}
            //else if (baskinlikSkorButon == 1)
            //{
            //    try
            //    {
            //        var saveFileDialog = new SaveFileDialog();
            //        saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            //        saveFileDialog.FilterIndex = 3;

            //        if (saveFileDialog.ShowDialog() == DialogResult.OK)
            //        {
            //            var workbook = new ExcelFile();
            //            var worksheet = workbook.Worksheets.Add("Sheet1");
            //            DataGridViewConverter.ImportFromDataGridView(worksheet, this.dgvTumBSkor, new ImportFromDataGridViewOptions() { ColumnHeaders = true });
            //            workbook.Save(saveFileDialog.FileName);
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Excel'e aktarma işlemi başarısız!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        return;
            //    }
            //}
        }
        //Karar Matrisi Kopyala-Yapıştır-Kes kodları ve kısayolları------------------------
        public void kararMatKopyalaYapistir()
        {

            try
            {
                DataGridViewRow selectedRow;
                /* Find first selected cell's row (or first selected row). */
                if (dataGridViewKararMat.SelectedRows.Count > 0)
                    selectedRow = dataGridViewKararMat.SelectedRows[0];
                else if (dataGridViewKararMat.SelectedCells.Count > 0)
                    selectedRow = dataGridViewKararMat.SelectedCells[0].OwningRow;
                else
                    return;
                /* Get clipboard Text */
                string clipText = Clipboard.GetText();
                /* Get Rows ( newline delimited ) */
                string[] rowLines = Regex.Split(clipText, "\r\n");
                foreach (string row in rowLines)
                {
                    /* Get Cell contents ( tab delimited ) */
                    string[] cells = Regex.Split(row, "\t");
                    DataGridViewRow r = new DataGridViewRow();
                    foreach (string sc in cells)
                    {
                        DataGridViewTextBoxCell c = new DataGridViewTextBoxCell();
                        c.Value = sc;
                        r.Cells.Add(c);
                    }
                    dataGridViewKararMat.Rows.Insert(selectedRow.Index, r);

                }


            }
            //catch (System.ArgumentException ex)
            //{

            //}
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void dataGridViewKararMat_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewKararMat.SelectedCells.Count > 0)
                dataGridViewKararMat.ContextMenuStrip = contextMenuStrip1;
            chkPasteToSelectedCells.Visible = true;
        }
        private void kesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //kopyalar
            CopyToClipboard();
            //seçilen hücreleri temizler
            if (tabControl1.SelectedTab == tabPageKararMatrisi)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewKararMat.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageNormalize)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewNormalize.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageBaskinlikSkorlari)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewBaskinlikSkoru.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageGenelBaskinlik)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewBaskinlikSkoru.SelectedCells)
                    dgvCell.Value = string.Empty;
            }

            else if (tabControl1.SelectedTab == tabPageAHP)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewKarsilastirmaMat.SelectedCells)
                    dgvCell.Value = string.Empty;
            }

            else if (tabControl1.SelectedTab == tabPageAHP)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewKriterAgirliklari.SelectedCells)
                    dgvCell.Value = string.Empty;
            }

            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewAyrintiKmat.SelectedCells)
                    dgvCell.Value = string.Empty;
            }

            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewC.SelectedCells)
                    dgvCell.Value = string.Empty;
            }

            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewWVektörü.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewDVektör.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewSonucKararMat.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewSonucNormalizeMat.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewSonucKriterAgirlik.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewSonucGoreliAgirlik.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewSonucKarsilastirmaMat.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                foreach (DataGridViewCell dgvCell in dataGridViewSonucGenelBaskinlik.SelectedCells)
                    dgvCell.Value = string.Empty;
            }

        }
        private void kopyalaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyToClipboard();
        }
        private void yapıştırToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PasteClipboardValue();
        }
        private void btnKararMatExceldeAc_Click(object sender, EventArgs e)
        {
            kararMatExcelAktar();
        }
        private void btnNormalizeEAc_Click(object sender, EventArgs e)
        {
            normalizeMatExcelAc();
        }
        private void baslangicToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageBaslangic;
        }
        private void yeniToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tumunuTemizle();
            tabControl1.SelectedTab = tabPageBaslangic;
        }
        private void exceleAktarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPageKararMatrisi)
            {
                btnKararMatEAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageNormalize)
            {
                btnNormalizeEAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageBaskinlikSkorlari)
            {
                btnBSkorEAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageGenelBaskinlik)
            {
                btnGBS_eAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageAHP)
            {
                btnIkiliKMatEAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                btnAhpKriterAğrBulEAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                btnSonucExcelAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageBaskinlikSkorSonuc)
            {
                btnTumKismiBskorExcelAktar.PerformClick();
            }
            else
            {
                MessageBox.Show("Excel'e aktarılabilecek bir matris yok", "Uyarı ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

        }
        private void cikisToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void kesToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPageKararMatrisi)
            {
                //kopyalar
                CopyToClipboard();
                //seçilen hücreleri temizler
                foreach (DataGridViewCell dgvCell in dataGridViewKararMat.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
            else if (tabControl1.SelectedTab == tabPageAgirlikBelirleme)
            {
                agirlikKopyala();

                foreach (DataGridViewCell dgvCell in dataGridViewAgirlik.SelectedCells)
                    dgvCell.Value = string.Empty;
            }
        }
        private void kopyalaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPageKararMatrisi)
            {
                CopyToClipboard();
            }
            else if (tabControl1.SelectedTab == tabPageAgirlikBelirleme)
            {
                agirlikKopyala();
            }
        }
        private void yapistirToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPageKararMatrisi)
            {
                PasteClipboardValue();
            }
            else if (tabControl1.SelectedTab == tabPageAgirlikBelirleme)
            {
                agirlikYapistir();
            }
        }
        private void kararMatrisiOlusturmaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (kriterler.Count > 1 && alternatifler.Count > 1)
            {
                tumunuTemizle();
                tiklanma = "olustur";
            }
            tabControl1.SelectedTab = tabPageKararMatOlusturma;

        }
        private void normalizeEdilmisKararMatrisiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageNormalize;
        }
        private void agirlikYontemSecToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageAgirlikBelirleme;
        }
        private void karsilastirmaMatrisiOlusturmaToolStripMenuItem_Click(object sender, EventArgs e)
        {

            tabControl1.SelectedTab = tabPageKarsilastirmaMat;
        }
        private void agirlikDegerleriToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageAHP;
        }
        private void ahpAyrintiliCozumToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageAhpAyrinti;
        }
        private void genelBaskinlikSkorlariToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageGenelBaskinlik;
        }
        private void kismiBaskinlikSkorlariToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            tumKismiBaskinlikSkorGoruntule();
            tabControl1.SelectedTab = tabPageBaskinlikSkorSonuc;
        }
        private void btnAKEkle_Click(object sender, EventArgs e)
        {
            if (tiklanma == "olustur" || tiklanma == "excel")
            {
                tabControl1.SelectedTab = tabPageKararMatOlusturma;
                if (listBoxKriter.Items.Count == 0)
                {
                    int i = 0;
                    foreach (var kriter in kriterler)
                    {
                        listBoxKriter.Items.Add(kriter + "  (" + faydaMaliyet[i].ToString() + ")");
                        i++;
                    }
                }
                if (listBoxAlternatif.Items.Count == 0)
                {
                    foreach (var alternatif in alternatifler)
                    {
                        listBoxAlternatif.Items.Add(alternatif);
                    }
                }

                //dataGridViewKararMat.AllowUserToAddRows = true;
            }
            else
            {
                MessageBox.Show("Eski çalışma için kriter ve alternatiflerde henüz değişiklik yapılamammaktadır!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

        }
        private void kismiBaskinlikSkoruAramaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageBaskinlikSkorlari;
        }
        private void btnAtla_Click(object sender, EventArgs e)
        {
            //if (tiklanma == "olustur")
            //{
            //    tabControl1.SelectedTab = tabPageKararMatOlusturma;
            //}
            //else if (tiklanma == "excel")
            //{

            //    tabControl1.SelectedTab = tabPageKararMatrisi;
            //}


        }
        private void dataGridViewNormalize_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewNormalize.SelectedCells.Count > 0)
                dataGridViewNormalize.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewBaskinlikSkoru_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewBaskinlikSkoru.SelectedCells.Count > 0)
                dataGridViewBaskinlikSkoru.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewGenelSkor_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewGenelSkor.SelectedCells.Count > 0)
                dataGridViewGenelSkor.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewKarsilastirmaMat_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewKarsilastirmaMat.SelectedCells.Count > 0)
                dataGridViewKarsilastirmaMat.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewKriterAgirliklari_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewKriterAgirliklari.SelectedCells.Count > 0)
                dataGridViewKriterAgirliklari.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewAyrintiKmat_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewAyrintiKmat.SelectedCells.Count > 0)
                dataGridViewAyrintiKmat.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewC_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewC.SelectedCells.Count > 0)
                dataGridViewC.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewWVektörü_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            if (dataGridViewWVektörü.SelectedCells.Count > 0)
                dataGridViewWVektörü.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewDVektör_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            if (dataGridViewDVektör.SelectedCells.Count > 0)
                dataGridViewDVektör.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewSonucKararMat_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewSonucKararMat.SelectedCells.Count > 0)
                dataGridViewSonucKararMat.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewSonucNormalizeMat_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewSonucNormalizeMat.SelectedCells.Count > 0)
                dataGridViewSonucNormalizeMat.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewSonucKriterAgirlik_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewSonucKriterAgirlik.SelectedCells.Count > 0)
                dataGridViewSonucKriterAgirlik.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewSonucGoreliAgirlik_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewSonucGoreliAgirlik.SelectedCells.Count > 0)
                dataGridViewSonucGoreliAgirlik.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewSonucKarsilastirmaMat_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewSonucKarsilastirmaMat.SelectedCells.Count > 0)
                dataGridViewSonucKarsilastirmaMat.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewSonucGenelBaskinlik_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewSonucGenelBaskinlik.SelectedCells.Count > 0)
                dataGridViewSonucGenelBaskinlik.ContextMenuStrip = contextMenuStrip1;
        }
        private void dataGridViewNormalize_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewBaskinlikSkoru_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewGenelSkor_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewKarsilastirmaMat_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewKriterAgirliklari_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewAyrintiKmat_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewC_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewWVektörü_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewDVektör_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewSonucKararMat_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewSonucNormalizeMat_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewSonucKriterAgirlik_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewSonucGoreliAgirlik_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewSonucKarsilastirmaMat_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void dataGridViewSonucGenelBaskinlik_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void btnOrnekExcelDosya_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm";
            saveFileDialog.DefaultExt = "xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                const string MyFileName = "NORMALİZE.xlsx"; //buraya indirmek istediğiniz dosyanın adını uzantısıyla birlikte yazın
                //daha sonra bu dosyayın bin deki debug klasörüne taşımamız gerekiyor
                string execPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
                var filePath = Path.Combine(execPath, MyFileName);
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook book = app.Workbooks.Open(filePath);
                book.SaveAs(saveFileDialog.FileName);
                book.Close();
                //buton üzerine gelindiğinde bilgi mesajı verdirelim

            }

        }
        private void kararMatrisiToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageKararMatrisi;
        }
        private void yeniToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            ileriButonlariniGizle();
            tabControl1.SelectedTab = tabPageBaslangic;
        }
        private void exceleAktarToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPageBaslangic || tabControl1.SelectedTab == tabPageKararMatOlusturma)
            {
                MessageBox.Show("Bu sayfada excel'e aktamak için bir matris bulunmuyor.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else if (tabControl1.SelectedTab == tabPageKararMatrisi)
            {
                btnKararMatEAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageNormalize)
            {
                btnNormalizeEAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageAHP)
            {
                btnIkiliKMatEAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                btnAhpKriterAğrBulEAktar.PerformClick();
            }

            else if (tabControl1.SelectedTab == tabPageGenelBaskinlik)
            {
                btnGBS_eAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageBaskinlikSkorSonuc)
            {
                btnTumKismiBskorExcelAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageBaskinlikSkorlari)
            {
                btnBSkorEAktar.PerformClick();
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                btnSonucExcelAktar.PerformClick();
            }

        }
        private void cikisToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }
        public void ileriButonlariniGorunurYap()
        {
            btnKararMatİleri.Visible = true;
            btnManuelİleri.Visible = true;
            btnNormalizeİleri.Visible = true;
            btnAhpİleri.Visible = true;
            btnAhpAyrintiİleri.Visible = true;
        }
        public void ileriButonlariniGizle()
        {
            btnKararMatİleri.Visible = false;
            btnManuelİleri.Visible = false;
            btnNormalizeİleri.Visible = false;
            btnAhpİleri.Visible = false;
            btnAhpAyrintiİleri.Visible = false;
        }
        public void eskiCalismaAc()
        {
            try
            {
                tumunuTemizle();
                ileriButonlariniGorunurYap();
                tiklanma = "eskiCalisma";
                verileriCagir();


            }
            catch (Exception ex)
            {
                MessageBox.Show("Dosya seçilemedi." + ex.Message, "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void btnEskiCalismaGetir_Click(object sender, EventArgs e)
        {
            tumunuTemizle();
            tiklanma = "eski";
            eskiCalismaAc();
        }
        public void agirlikBelirleEAktar()
        {
            try
            {
                ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Worksheets1");
                OfficeOpenXml.ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                worksheet.Cells[1, 1].Value = "Kriterlere Ait Ağırlık Değerleri";
                worksheet.Cells[1, 1, 1, 4].Merge = true;

                var columns = dataGridViewAgirlik.Columns;
                for (int i = 0; i < columns.Count; i++)
                {
                    worksheet.Cells[2, i + 1].Value = columns[i].HeaderText;
                }

                int rowIndex = 3;
                var rows = dataGridViewAgirlik.Rows;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].Cells[0] != null)
                    {
                        for (int j = 0; j < rows[i].Cells.Count; j++)
                        {
                            worksheet.Cells[rowIndex, j + 1].Value = rows[i].Cells[j].Value;
                        }
                        rowIndex++;
                    }
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Dosyası|*.xlsx";
                saveFileDialog.ShowDialog();

                Stream stream = saveFileDialog.OpenFile();
                package.SaveAs(stream);
                stream.Close();

                MessageBox.Show("Excel dosyanız başarıyla kaydedildi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Excel'e aktarma işlemi başarısız.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridViewSonucGoreliAgirlik_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void panel27_Paint(object sender, PaintEventArgs e)
        {

        }
        private void açToolStripMenuItem_Click(object sender, EventArgs e)
        {
            eskiCalismaAc();
        }
        private void kaydetToolStripMenuItem_Click(object sender, EventArgs e)
        {//çalışmanın yolunu ve adını tut o yoldaki o addaki excel üzerine kaydettir
            matrisleriKaydet();
        }
        private void farklıKaydetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            matrisleriKaydet();
        }
        private void btnKararMatİleri_Click(object sender, EventArgs e)
        {
            if (secilenNormalizeYontemi == "1")
            {

                normalizeEdilmisKararMatrisiToolStripMenuItem.Visible = true;
                normalizeMatCerceve();
                //boş hücre kontrolü
                for (int i = 0; i < alternatifler.Count; i++) //satır
                {
                    for (int j = 1; j < kriterler.Count + 1; j++) //sutun
                    {
                        if (dataGridViewKararMat.Rows[i].Cells[j].Value == null)
                        {
                            MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                    }
                }
                minMaxNormalization();

                if (tiklanma == "eski")
                {
                    pnlYontemSec.Visible = true;
                }

                tabControl1.SelectedTab = tabPageNormalize;

            }
            else if (secilenNormalizeYontemi == "2")
            {
                normalizeEdilmisKararMatrisiToolStripMenuItem.Visible = true;
                normalizeMatCerceve();
                for (int i = 0; i < alternatifler.Count; i++) //satır
                {
                    for (int j = 1; j < kriterler.Count + 1; j++) //sutun
                    {
                        if (dataGridViewKararMat.Rows[i].Cells[j].Value == null)
                        {
                            MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                normalizeYeni();

                if (tiklanma == "eski")
                {
                    pnlYontemSec.Visible = true;
                }
            }
        }
        private void btnNormalizeİleri_Click(object sender, EventArgs e)
        {
            if (pnlYontemSec.Visible == false)
            {
                pnlYontemSec.Visible = true;

            }

            if (yontem == "ahp")
            {
                tabControl1.SelectedTab = tabPageAHP;
            }
            else if (yontem == "manuel")
            {
                tabControl1.SelectedTab = tabPageAgirlikBelirleme;
            }

            //eğer normalize yöntemi değiştirildiyse ve normalize matrisinin bulunduğu sayfadaki ileri butonuna basıldıysa- 
            //yeni normalize yöntemi ve eski ağırlık değerleri kulllanılarak yeniden hesaplama yaptırılır




        }
        private void btnAgirlikYontemİleri_Click(object sender, EventArgs e)
        {
            if (yontem == "ahp")
            {
                tabControl1.SelectedTab = tabPageAHP;
            }
            else if (yontem == "manuel")
            {
                goreliAgirliklarim();
                maxGenelBakinlik();
                genelBaskinlikSkoruMatrisi();
                tabControl1.SelectedTab = tabPageGenelBaskinlik;
            }
        }
        private void btnAhpİleri_Click(object sender, EventArgs e)
        {
            goreliAgirliklarim();
            maxGenelBakinlik();
            genelBaskinlikSkoruMatrisi();

            //SONUÇLAR
            sonucKararMatDoldur();
            sonucNormalizeMatDoldur();
            sonucKriterAgirlikDoldur();
            sonucGoreliAgirlikDoldur();
            sonucGenelBaskinlikDoldur();
            if (yontem == "ahp")
            {
                sonucKarsilastirmaMatDoldur();
            }

            //tüm kısmi baskınlık skoru görüntüleme
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            tumKismiBaskinlikSkorGoruntule();
            tabControl1.SelectedTab = tabPageGenelBaskinlik;


            //if (secilenNormalizeYontemi != normalizeYonSakla)
            //{
            //    goreliAgirliklarim();
            //    maxGenelBakinlik();
            //    genelBaskinlikSkoruMatrisi();

            //    //SONUÇLAR
            //    sonucKararMatDoldur();
            //    sonucNormalizeMatDoldur();
            //    sonucKriterAgirlikDoldur();
            //    sonucGoreliAgirlikDoldur();
            //    sonucGenelBaskinlikDoldur();
            //    if (yontem == "ahp")
            //    {
            //        sonucKarsilastirmaMatDoldur();
            //    }

            //    //tüm kısmi baskınlık skoru görüntüleme
            //    flowLayoutPanel1.Controls.Clear();
            //    flowLayoutPanel3.Controls.Clear();
            //    tumKismiBaskinlikSkorGoruntule();
            //    tabControl1.SelectedTab = tabPageGenelBaskinlik;

            //}
            //else if (secilenNormalizeYontemi == normalizeYonSakla)
            //{
            //    tabControl1.SelectedTab = tabPageGenelBaskinlik;
            //}


        }
        private void kararMatrisiToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void groupBox1_Move(object sender, EventArgs e)
        {

        }
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        private void btnAgirlikExcelAl_Click(object sender, EventArgs e)
        {
            agirliklariExceldenAl();
        }
        private void btnAhpAyrintiİleri_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageGenelBaskinlik;
        }
        private void btnAgirlikBelirleEAktar_Click(object sender, EventArgs e)
        {
            agirlikBelirleEAktar();
        }
        private void btnTamam_Click(object sender, EventArgs e)
        {

            //if (txtCalismaAd.Text != "")
            //{
            //    if (tiklanma == "olustur")
            //    {
            //        calismaIsmi = txtCalismaAd.Text;
            //        tabControl1.SelectedTab = tabPageKararMatOlusturma;
            //    }
            //    else if (tiklanma == "excel")
            //    {
            //        calismaIsmi = txtCalismaAd.Text;
            //        tabControl1.SelectedTab = tabPageKararMatrisi;
            //    }

            //}
            //else
            //{
            //    MessageBox.Show("Lütfen bir metin giriniz!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;

            //}

        }
        private void hakkindaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageHakkinda;
        }
        private void ayrintiliCozumToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPageSonuc;
        }
        private void btnBSkorEAc_Click(object sender, EventArgs e)
        {
            baskinlikSkorExcelAc();
        }
        private void btnAzalanSira_Click(object sender, EventArgs e)
        {
            dataGridViewGenelSkor.Sort(dataGridViewGenelSkor.Columns[2], ListSortDirection.Descending);//Normal Sıralama
        }
        private void btnArtanSira_Click(object sender, EventArgs e)
        {
            dataGridViewGenelSkor.Sort(dataGridViewGenelSkor.Columns[2], ListSortDirection.Ascending);//Normal Sıralama

        }
        private void lblTetaDeğiştir_Click(object sender, EventArgs e)
        {
            if (btnTetaKaydet.Visible == false)
            {
                btnTetaKaydet.Visible = true;
                txtTeta.Visible = true;
            }
            else if (btnTetaKaydet.Visible == true)
            {
                btnTetaKaydet.Visible = false;
                txtTeta.Visible = false;
            }
        }
        private void btnTetaKaydet_Click(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToDouble(txtTeta.Text) >= 0 && Convert.ToDouble(txtTeta.Text) <= 1)
                {
                    dataGridViewGenelSkor.Rows.Clear();
                    θ = Convert.ToDouble(txtTeta.Text);
                    goreliAgirliklarim();
                    maxGenelBakinlik();
                    genelBaskinlikSkoruMatrisi();
                }
                else
                {
                    MessageBox.Show("Lütfen 0-1 arasında bir değer giriniz!", "UYARI ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Baskınlık skorları hesaplanamadı!", "HATA ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        private void txtTeta_Enter(object sender, EventArgs e)
        {
            txtTeta.Text = "";
        }
        private void btnIkiliKMatEAc_Click(object sender, EventArgs e)
        {
            ikiliKMatExcelAc();
        }
        private void btnCalismayiKaydet_Click(object sender, EventArgs e)
        {
            matrisleriKaydet();
        }
        private void dataGridViewKararMat_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            CopyToClipboard();
                            break;

                        case Keys.V:
                            PasteClipboardValue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void CopyToClipboard()
        {
            if (tabControl1.SelectedTab == tabPageKararMatrisi)
            {
                DataObject dataObj = dataGridViewKararMat.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageNormalize)
            {
                DataObject dataObj = dataGridViewNormalize.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageBaskinlikSkorlari)
            {
                DataObject dataObj = dataGridViewBaskinlikSkoru.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageGenelBaskinlik)
            {

                DataObject dataObj = dataGridViewBaskinlikSkoru.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageAHP)
            {
                DataObject dataObj = dataGridViewKarsilastirmaMat.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageAHP)
            {
                DataObject dataObj = dataGridViewKriterAgirliklari.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                DataObject dataObj = dataGridViewAyrintiKmat.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                DataObject dataObj = dataGridViewC.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                DataObject dataObj = dataGridViewWVektörü.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            {
                DataObject dataObj = dataGridViewDVektör.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                DataObject dataObj = dataGridViewSonucKararMat.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                DataObject dataObj = dataGridViewSonucNormalizeMat.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                DataObject dataObj = dataGridViewSonucKriterAgirlik.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                DataObject dataObj = dataGridViewSonucGoreliAgirlik.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                DataObject dataObj = dataGridViewSonucKarsilastirmaMat.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            else if (tabControl1.SelectedTab == tabPageSonuc)
            {
                DataObject dataObj = dataGridViewSonucGenelBaskinlik.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
        }
        private void PasteClipboardValue() //pano değerlerini yapıştır
        {
            if (tabControl1.SelectedTab == tabPageKararMatrisi)
            {
                //hiçbir hücre seçilmezse
                if (dataGridViewKararMat.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewKararMat);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewKararMat.Columns.Count - 1
                        && iRowIndex <= dataGridViewKararMat.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewKararMat[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }


            }

            else if (tabControl1.SelectedTab == tabPageNormalize)
            { //hiçbir hücre seçilmezse
                if (dataGridViewNormalize.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewNormalize);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewNormalize.Columns.Count - 1
                        && iRowIndex <= dataGridViewNormalize.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewNormalize[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageBaskinlikSkorlari)
            { //hiçbir hücre seçilmezse
                if (dataGridViewBaskinlikSkoru.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewBaskinlikSkoru);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewBaskinlikSkoru.Columns.Count - 1
                        && iRowIndex <= dataGridViewBaskinlikSkoru.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewBaskinlikSkoru[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageGenelBaskinlik)
            { //hiçbir hücre seçilmezse
                if (dataGridViewBaskinlikSkoru.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewBaskinlikSkoru);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewBaskinlikSkoru.Columns.Count - 1
                        && iRowIndex <= dataGridViewBaskinlikSkoru.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewBaskinlikSkoru[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }




            }

            else if (tabControl1.SelectedTab == tabPageAHP)
            { //hiçbir hücre seçilmezse
                if (dataGridViewKarsilastirmaMat.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewKarsilastirmaMat);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewKarsilastirmaMat.Columns.Count - 1
                        && iRowIndex <= dataGridViewKarsilastirmaMat.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewKarsilastirmaMat[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageAHP)
            { //hiçbir hücre seçilmezse
                if (dataGridViewKriterAgirliklari.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewKriterAgirliklari);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewKriterAgirliklari.Columns.Count - 1
                        && iRowIndex <= dataGridViewKriterAgirliklari.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewKriterAgirliklari[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            { //hiçbir hücre seçilmezse
                if (dataGridViewAyrintiKmat.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewAyrintiKmat);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewAyrintiKmat.Columns.Count - 1
                        && iRowIndex <= dataGridViewAyrintiKmat.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewAyrintiKmat[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            { //hiçbir hücre seçilmezse
                if (dataGridViewC.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewC);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewC.Columns.Count - 1
                        && iRowIndex <= dataGridViewC.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewC[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            { //hiçbir hücre seçilmezse
                if (dataGridViewWVektörü.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewWVektörü);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewWVektörü.Columns.Count - 1
                        && iRowIndex <= dataGridViewWVektörü.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewWVektörü[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageAhpAyrinti)
            { //hiçbir hücre seçilmezse
                if (dataGridViewDVektör.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewDVektör);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewDVektör.Columns.Count - 1
                        && iRowIndex <= dataGridViewDVektör.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewDVektör[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageSonuc)
            { //hiçbir hücre seçilmezse
                if (dataGridViewSonucKararMat.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewSonucKararMat);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewSonucKararMat.Columns.Count - 1
                        && iRowIndex <= dataGridViewSonucKararMat.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewSonucKararMat[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageSonuc)
            { //hiçbir hücre seçilmezse
                if (dataGridViewSonucNormalizeMat.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewSonucNormalizeMat);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewSonucNormalizeMat.Columns.Count - 1
                        && iRowIndex <= dataGridViewSonucNormalizeMat.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewSonucNormalizeMat[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageSonuc)
            { //hiçbir hücre seçilmezse
                if (dataGridViewSonucKriterAgirlik.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewSonucKriterAgirlik);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewSonucKriterAgirlik.Columns.Count - 1
                        && iRowIndex <= dataGridViewSonucKriterAgirlik.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewSonucKriterAgirlik[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageSonuc)
            { //hiçbir hücre seçilmezse
                if (dataGridViewSonucGoreliAgirlik.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewSonucGoreliAgirlik);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewSonucGoreliAgirlik.Columns.Count - 1
                        && iRowIndex <= dataGridViewSonucGoreliAgirlik.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewSonucGoreliAgirlik[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageSonuc)
            { //hiçbir hücre seçilmezse
                if (dataGridViewSonucKarsilastirmaMat.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewSonucKarsilastirmaMat);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewSonucKarsilastirmaMat.Columns.Count - 1
                        && iRowIndex <= dataGridViewSonucKarsilastirmaMat.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewSonucKarsilastirmaMat[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

            else if (tabControl1.SelectedTab == tabPageSonuc)
            { //hiçbir hücre seçilmezse
                if (dataGridViewSonucGenelBaskinlik.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //başlangıç hücresini alma
                DataGridViewCell startCell = GetStartCell(dataGridViewSonucGenelBaskinlik);
                //pano değerlerini sözlükten alma
                Dictionary<int, Dictionary<int, string>> cbValue =
                        ClipBoardValues(Clipboard.GetText());

                int iRowIndex = startCell.RowIndex;
                foreach (int rowKey in cbValue.Keys)
                {
                    int iColIndex = startCell.ColumnIndex;
                    foreach (int cellKey in cbValue[rowKey].Keys)
                    {
                        //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                        if (iColIndex <= dataGridViewSonucGenelBaskinlik.Columns.Count - 1
                        && iRowIndex <= dataGridViewSonucGenelBaskinlik.Rows.Count - 1)
                        {
                            DataGridViewCell cell = dataGridViewSonucGenelBaskinlik[iColIndex, iRowIndex];

                            // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                            if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                                (!chkPasteToSelectedCells.Checked))
                                cell.Value = cbValue[rowKey][cellKey];
                        }
                        iColIndex++;
                    }
                    iRowIndex++;
                }



            }

        }
        private DataGridViewCell GetStartCell(DataGridView dgView)
        {

            // en küçük satırı, sütun dizinini al
            if (dgView.SelectedCells.Count == 0)
                return null;

            int rowIndex = dgView.Rows.Count - 1;
            int colIndex = dgView.Columns.Count - 1;

            foreach (DataGridViewCell dgvCell in dgView.SelectedCells)
            {
                if (dgvCell.RowIndex < rowIndex)
                    rowIndex = dgvCell.RowIndex;
                if (dgvCell.ColumnIndex < colIndex)
                    colIndex = dgvCell.ColumnIndex;
            }

            return dgView[colIndex, rowIndex];
        }
        private Dictionary<int, Dictionary<int, string>> ClipBoardValues(string clipboardValue)
        {
            Dictionary<int, Dictionary<int, string>>
            copyValues = new Dictionary<int, Dictionary<int, string>>();

            String[] lines = clipboardValue.Split('\n');

            for (int i = 0; i <= lines.Length - 1; i++)
            {
                copyValues[i] = new Dictionary<int, string>();
                String[] lineContent = lines[i].Split('\t');

                // boş bir hücre değeri kopyalandıysa, sözlüğü boş bir dize ile ayarlama
                if (lineContent.Length == 0)
                    copyValues[i][0] = string.Empty;
                else
                {
                    for (int j = 0; j <= lineContent.Length - 1; j++)
                        copyValues[i][j] = lineContent[j];
                }
            }
            return copyValues;
        }
        //--------------------------------------------------------------    
        public void kararMatImportListeDoldurma()
        {
            try
            {
                kriterler.Clear();
                faydaMaliyet.Clear();
                alternatifler.Clear();
                for (int i = 1; i < dataGridViewKararMat.Columns.Count; i++)
                {
                    kriterler.Add(dataGridViewKararMat.Columns[i].Name.ToString());
                }
                for (int i = 0; i < kriterler.Count; i++)
                {
                    faydaMaliyet.Add(rbtnFayda.Text);
                }
                for (int i = 0; i < dataGridViewKararMat.Rows.Count; i++)
                {
                    alternatifler.Add(dataGridViewKararMat.Rows[i].Cells[0].Value.ToString());
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        public void kararMatListeDoldurmaEskiCalisma()
        {
            try
            {
                kriterler.Clear();

                alternatifler.Clear();
                for (int i = 1; i < dataGridViewKararMat.Columns.Count; i++)
                {
                    kriterler.Add(dataGridViewKararMat.Columns[i].Name.ToString());
                }

                for (int i = 0; i < dataGridViewKararMat.Rows.Count; i++)
                {
                    alternatifler.Add(dataGridViewKararMat.Rows[i].Cells[0].Value.ToString());
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Lütfen boş hücreleri doldurunuz!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        public void importKararMatDoldur()
        {
            gridTasarimSirasiz(dataGridViewImport);
            try
            {
                dataGridViewKararMat.Columns.Clear();
                dataGridViewKararMat.Rows.Clear();

                dataGridViewKararMat.ColumnCount = dataGridViewImport.Columns.Count;
                dataGridViewKararMat.Columns[0].Name = "Virgül ile ayırınız (nokta kullanmayınız)";

                for (int i = 1; i < dataGridViewImport.Columns.Count; i++)
                {
                    dataGridViewKararMat.Columns[i].Name = (dataGridViewImport.Rows[0].Cells[i].Value.ToString());
                }

                for (int i = 1; i < dataGridViewImport.Rows.Count; i++)
                {
                    dataGridViewKararMat.Rows.Add(dataGridViewImport.Rows[i].Cells[0].Value.ToString());
                }

                for (int i = 1; i < dataGridViewImport.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridViewImport.Columns.Count; j++)
                    {

                        dataGridViewKararMat.Rows[i - 1].Cells[j].Value = dataGridViewImport.Rows[i].Cells[j].Value.ToString();


                    }
                }

                //İLK SUTUNDAKİ DEĞERLERİN DEĞİŞTİRİLMESİNİ ENGELLEDİM
                for (int rC = 0; rC < alternatifler.Count; rC++)
                {
                    dataGridViewKararMat.Rows[rC].Cells[0].ReadOnly = true;
                }

                dataGridViewImport.Visible = false;
            }
            catch (Exception ex)
            {

                MessageBox.Show("Hata!" + ex.Message, "Hata ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        //Ağırlık Matrisi Kopyala-Yapıştır-Kes kodları ve kısayolları------------------------
        public void agirlikMatKopyalaYapistir()
        {

            try
            {
                DataGridViewRow selectedRow;
                /* Find first selected cell's row (or first selected row). */
                if (dataGridViewAgirlik.SelectedRows.Count > 0)
                    selectedRow = dataGridViewAgirlik.SelectedRows[0];
                else if (dataGridViewAgirlik.SelectedCells.Count > 0)
                    selectedRow = dataGridViewAgirlik.SelectedCells[0].OwningRow;
                else
                    return;
                /* Get clipboard Text */
                string clipText = Clipboard.GetText();
                /* Get Rows ( newline delimited ) */
                string[] rowLines = Regex.Split(clipText, "\r\n");
                foreach (string row in rowLines)
                {
                    /* Get Cell contents ( tab delimited ) */
                    string[] cells = Regex.Split(row, "\t");
                    DataGridViewRow r = new DataGridViewRow();
                    foreach (string sc in cells)
                    {
                        DataGridViewTextBoxCell c = new DataGridViewTextBoxCell();
                        c.Value = sc;
                        r.Cells.Add(c);
                    }
                    dataGridViewAgirlik.Rows.Insert(selectedRow.Index, r);

                }


            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void dataGridViewAgirlik_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridViewAgirlik.SelectedCells.Count > 0)
                dataGridViewAgirlik.ContextMenuStrip = contextMenuStripAgirlik;
            //chkPasteToSelectedCells.Visible = true;
        }
        private void toolStripMenuItemAgirlikKes_Click(object sender, EventArgs e)
        {
            agirlikKopyala();

            foreach (DataGridViewCell dgvCell in dataGridViewAgirlik.SelectedCells)
                dgvCell.Value = string.Empty;
        }
        private void toolStripMenuItemAgirlikKopyala_Click(object sender, EventArgs e)
        {
            agirlikKopyala();

        }
        private void toolStripMenuItemAgirlikYapistir_Click(object sender, EventArgs e)
        {
            agirlikYapistir();
        }
        private void dataGridViewAgirlik_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            agirlikKopyala();
                            break;

                        case Keys.V:
                            agirlikYapistir();
                            break;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kopyalama / yapıştırma işlemi başarısız oldu." + ex.Message, "Kopyala/Yapıştır", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void agirlikKopyala()
        {
            DataObject dataObj = dataGridViewAgirlik.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void agirlikYapistir() //pano değerlerini yapıştır
        {
            //hiçbir hücre seçilmezse
            if (dataGridViewAgirlik.SelectedCells.Count == 0)
            {
                MessageBox.Show("Lütfen bir hücre seçin", "Yapıştır",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //başlangıç hücresini alma
            DataGridViewCell startCell = GetStartCellAgirlik(dataGridViewAgirlik);
            //pano değerlerini sözlükten alma
            Dictionary<int, Dictionary<int, string>> cbValue =
                    ClipBoardValues(Clipboard.GetText());

            int iRowIndex = startCell.RowIndex;
            foreach (int rowKey in cbValue.Keys)
            {
                int iColIndex = startCell.ColumnIndex;
                foreach (int cellKey in cbValue[rowKey].Keys)
                {
                    //dizinin sınırlar dahilinde olup olmadığını kontrol etme
                    if (iColIndex <= dataGridViewAgirlik.Columns.Count - 1
                    && iRowIndex <= dataGridViewAgirlik.Rows.Count - 1)
                    {
                        DataGridViewCell cell = dataGridViewAgirlik[iColIndex, iRowIndex];

                        // 'chkPasteToSelectedCells' işaretliyse seçili hücrelere kopyala
                        if ((chkPasteToSelectedCells.Checked && cell.Selected) ||
                            (!chkPasteToSelectedCells.Checked))
                            cell.Value = cbValue[rowKey][cellKey];
                    }
                    iColIndex++;
                }
                iRowIndex++;
            }
        }
        private DataGridViewCell GetStartCellAgirlik(DataGridView dgView)
        {

            // en küçük satırı, sütun dizinini al
            if (dgView.SelectedCells.Count == 0)
                return null;

            int rowIndex = dgView.Rows.Count - 1;
            int colIndex = dgView.Columns.Count - 1;

            foreach (DataGridViewCell dgvCell in dgView.SelectedCells)
            {
                if (dgvCell.RowIndex < rowIndex)
                    rowIndex = dgvCell.RowIndex;
                if (dgvCell.ColumnIndex < colIndex)
                    colIndex = dgvCell.ColumnIndex;
            }

            return dgView[colIndex, rowIndex];
        }
        private Dictionary<int, Dictionary<int, string>> ClipBoardValuesAgirlik(string clipboardValue)
        {
            Dictionary<int, Dictionary<int, string>>
            copyValues = new Dictionary<int, Dictionary<int, string>>();

            String[] lines = clipboardValue.Split('\n');

            for (int i = 0; i <= lines.Length - 1; i++)
            {
                copyValues[i] = new Dictionary<int, string>();
                String[] lineContent = lines[i].Split('\t');

                // boş bir hücre değeri kopyalandıysa, sözlüğü boş bir dize ile ayarlayın
                // else Değeri sözlüğe ayarla
                if (lineContent.Length == 0)
                    copyValues[i][0] = string.Empty;
                else
                {
                    for (int j = 0; j <= lineContent.Length - 1; j++)
                        copyValues[i][j] = lineContent[j];
                }
            }
            return copyValues;
        }
        //--------------------------------------------------------------    


    }
}
