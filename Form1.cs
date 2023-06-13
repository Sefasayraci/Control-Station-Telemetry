using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Reflection.Emit;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;
using offis = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using GMap.NET.WindowsForms;
using GMap.NET;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms.Markers;
using System.Reflection;
using System.Linq.Expressions;
using System.Runtime.ConstrainedExecution;


//-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!PORTA KART DEVRESİNİ TAKTIĞIMIZDA YER İSTASYONUNU AKTİFLEŞTİRMEDEN 5 SN BOYUNCA BEKLE YA DA TEKRAR GİR VE AYNI ŞEKİLDE DENE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
//  TEKNİK KODLARI KONTROL ET. UÇUŞ KONTROL YAZILIMI İÇİN.
// MAP'TEN DOLAYI APPLİCATİON AÇILMIYOR O DÜZELECEK.
// PORT SEÇME SEÇENEĞİ EKLENECEK
// PORT BAĞLANTI AÇIK KAPALI DURUMUNDA DEFAULT BİR ŞEYLER EKLENEBİLİR.

// TASARIM KISMINDA DEĞİŞİKLİK YAPILDI. DAİRE FOTOĞRAF EKLENDİ.
// RESİMDEN BUTON ÜRETİLEBİLİR.
// EXCEL BUTONLARI YERİNE IMAGE E GEÇİLDİ.
// timer_tick'te verilerim 1 sn interval ile gelmektedir.
// Panel ekran boyutu ayarlanacak
// HEM BURADA HEM TASARIMDA YAZILAR DEĞŞECEK
// ANA YAZILIMLA KODLAR KARŞILAŞTIRILACAK
// EKRAN BOYUTU 1920X1080 DEĞİL 1100 AYARLADIM VERİ KAYDETMEK İÇİN
// LOKAL DOSYA AÇMA DÜZELTİLDİ.
// FORM BORDER STYLE KAPATIIM, İLERİ Kİ ZAMANDA AÇABİLİRİM - IMAGE BUTTON EKLENİP DEĞİŞTİRİLEBİLİR. 
// GÖREV YÜKÜ İÇİN GPS VERİLER KOYULACAK
// 5 HZ DE BİR VERİ İSTENİYOR *5 Hz frekansla (her farklı veri grubundan saniyede 5 veri yayımlanması) yer istasyonuna iletmesi gerekmektedir.* ÖNEMLİ TEPE NOKTASINDAN İTİBAREN
// YER İSTASYONU FTD Dİ İLE PAKET GÖNDERME SEÇENEĞİ OLACAK
// GÖREV YÜKÜNDE BASINÇ SENSÖRÜ KULLANILMASI DURUMUDA GERÇEKLEŞCEKTİR. BUNUN İÇİN YER İSTASYONUNU AYARLA
// BUZZER TEXTBOX'A YERLEŞECEK.
// NEM VERİSİ EKLENECEK.
// NEM VERİSİ HANGİ SENSÖRDEN NASIL BULUNUYOR
// 2 ADET I2C DEN VERİ OKUMA DA SIKINTI ÇIKACAK MI SON KEZ KONTROL ET.
// BUZZER AYRICA UÇUŞ YAZILIMINDA SENSÖRLERİN TESTLERİ İÇİNDE KULLANILABİLİR.
// TAKIM ID Sİ GÖNDER
// AYRI BİR PORT AÇ VE YER İSTASYONUNA GÖNDER HAKEM YER İSTASYONUNA 11520 OLACAK
// HER VERİ GİTTİĞİNDE SAYACA +1 EKLENECEK BUNUDA LABEL KOY
// ARAYÜZ LOORA İLE HABERLEŞMESİ GERÇEKLEŞTİ VE COM SEÇNEĞİDE DEĞİŞTİRİLDİ. ARTIK BUNA GÖRE TASARIM YAPILIYOR. AYRICA PORT SEÇNEĞİNİN YAZILIMINI BAKABİLİRSİN. VERİ ALMADA MPU6050 SENSÖR EKLENDİ.
// AÇI VERİSİ EKLENECEK.
// NEM SENSÖR VERİSİ OLACAK MI?
// TAKIM ID VE PAKET SAYISI İÇİN KONTROL PANELİNE VERİ İÇİN COMBOBOX EKLENİP YAZILACAK. VE BU ID UÇUŞ KODUNA EKLENECEK.
// ALTTAKİ BEYAZ CHART KONTROL EDİLECEK
//-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------




namespace tulpar_ground
{
    public partial class tulpar_form : Form
    {

        string sonuc;

        long maksmx = 30, minmx = 0;
        //  long maksmy = 30, minmy = 0;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Burada Excel'e veri kaydetmek için zaman, satır ve sütun fonksiyonlarını tanıttım. Ayrıca veri akışında anlık olarak tarih ve saat verisi anlık olarak aktarılacak C# fonksiyon kütüphanesi olarak 'Date' olarak tanıtılmıştır.
        DateTime yeni = DateTime.Now;
        int zaman = 0;
        int satir = 1;
        int sutun = 1;
        int satirNo = 1;
        int k = 0;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        double enlem, boylam = 0;
        bool chk;


        public tulpar_form()
        {

             InitializeComponent();
             serialPort1.PortName = "COM8";
             serialPort1.BaudRate = 9600;
            
            // KÜTÜPHANELER BİTTİ GERİYE SADECE KODU DÜZENLEMESİ KALDI YARISI BOŞTA ŞU AN
            // LİNK: https://www.youtube.com/watch?v=WpfjRaYVId8
            // LİNK: https://www.youtube.com/watch?v=TxSJJfaAzKg&t=3s

            //*****************GMAP OLAYLARI*************************
            //  map.MapProvider = GMapProviders.GoogleMap;
            gMap_Nav.DragButton = MouseButtons.Left;
            gMap_Nav.MapProvider = GMapProviders.GoogleMap;
            double lat = 39.940396;                    // Ankara Gazi Üniversitesi Enlem:39.93765208023499 Boylam:32.8214810699122. BURAYA DA BU KONUMLAR YAZILACAK, MANİSA ŞİMDİLİK GEÇİCİ
            double lng = 32.819424;                   // Manisa Enlem: 38.616890  Manisa Boylam: 27.426498
          //  gMap_Nav.Position = new GMap.NET.PointLatLng(39.940396, 32.819424); // BURAYA TELEMETRİ VERSİSİ OLAN GPS ENLEM-BOYLAM OLARAK YAZILACAK VE HER KUTUSUNUN İSMİ DÜZELTİLECEKTİR.
            gMap_Nav.MinZoom = 5;
            gMap_Nav.MaxZoom = 100;
            gMap_Nav.Zoom = 19; // DEFAULT AYARI 10 OLARAK BAŞLADIM
            
            GMap.NET.PointLatLng point = new GMap.NET.PointLatLng(lat, lng);
            GMapMarker marker = new GMarkerGoogle(point, GMarkerGoogleType.blue_dot);
          
            // Create a Overlay
            GMapOverlay markers = new GMapOverlay("markers");
          
            //Add all available markers to that Overlay
            markers.Markers.Add(marker);
          
            //Overlap map with Overlay
            gMap_Nav.Overlays.Add(markers);

            //*******************GMAP OLAYLARI******************************

            // KÜTÜPHANELER BİTTİ GERİYE SADECE KODU DÜZENLEMESİ KALDI YARISI BOŞTA ŞU AN
            // LİNK: https://www.youtube.com/watch?v=WpfjRaYVId8
            // LİNK: https://www.youtube.com/watch?v=TxSJJfaAzKg&t=3s
            
            //*****************GMAP2 OLAYLARI*************************
            //  map.MapProvider = GMapProviders.GoogleMap;
            gMap_Nav2.DragButton = MouseButtons.Left;
            gMap_Nav2.MapProvider = GMapProviders.GoogleMap;
            double lat2 = 39.91587892502433;                 // Ankara Gazi Üniversitesi Enlem:39.93765208023499 Boylam:32.8214810699122. BURAYA DA BU KONUMLAR YAZILACAK, MANİSA ŞİMDİLİK GEÇİCİ
            double lng2 = 32.82822954904327;
          //  gMap_Nav2.Position = new GMap.NET.PointLatLng(39.91587892502433, 32.82822954904327); // BURAYA TELEMETRİ VERSİSİ OLAN GPS ENLEM-BOYLAM OLARAK YAZILACAK VE HER KUTUSUNUN İSMİ DÜZELTİLECEKTİR.
            gMap_Nav2.MinZoom = 5;
            gMap_Nav2.MaxZoom = 100;
            gMap_Nav2.Zoom = 19; // DEFAULT AYARI 10 OLARAK BAŞLADIM

            GMap.NET.PointLatLng point2 = new GMap.NET.PointLatLng(lat, lng);
            GMapMarker marker2 = new GMarkerGoogle(point, GMarkerGoogleType.blue_dot);

            // Create a Overlay
            GMapOverlay markers2 = new GMapOverlay("markers");

            //Add all available markers to that Overlay
            markers.Markers.Add(marker);

            //Overlap map with Overlay
            gMap_Nav2.Overlays.Add(markers);

            //*******************GMAP2 OLAYLARI******************************
            


        }

        private void tulpar_form_Load(object sender, EventArgs e)
        {/*
            //SerialPort serialPort = new SerialPort(); 
            foreach (string port in portlar)
            {
                comboBox1.Items.Add(port);
                comboBox1.SelectedIndex = 0;
                label2.Text = "Bağlantı Kuruldu";
            }
            serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);*/
        }
        
        private void baglanti_Click(object sender, EventArgs e)                                             // PORT SEÇME SEÇENEĞİ EKLENECEK
        {

            try
            {

                serialPort1.Open();
                timer1.Start();
                baglanti.Enabled = false;                         // BUNU SİLME NEDENİM BAĞLANTİ KES'TENSONRA TEKRAR BASILMASI DURUMUDUR.

                  baglanti_kontrol.Text = "Port Bağlantısı Gerçekleşti";
                  baglanti_kontrol.ForeColor = Color.Green;

            }
            catch (Exception)
            {
                MessageBox.Show("Port Bilgisi Alınamadı", "Bağlantı Hatası", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Warning); // Durdur, Yeniden Dene, Yoksay
                // await Task.Delay(1000);
            } // https://mustafabukulmez.com/2018/01/20/c-messagebox-isinize-yarayacak-tum-ozellikleri/   // KOD İK ÇALIŞTIĞINDA BAĞLANTI KURULMADI OLARAK TEXTBOX BUNİFU'YA YAZILACAK

        }

        private void baglanti_kes_Click(object sender, EventArgs e)
        {

            try
            {

                serialPort1.Close();
                timer1.Stop();
                baglanti_kes.Enabled = true;
                baglanti.Enabled = true;            // BUNU EKLEME NEDENİM BUTON TEKAR AKTİF OLSUN VE BASILDIĞINDA TEKRAR BASILIRSA KOD PATLAMASIN VERİ ALMA İŞLEMİ DEVAM ETSİN DİYE YAPTIM.


                 baglanti_kontrol.Text = "Port Bağlantısı Kesildi";
                 baglanti_kontrol.ForeColor = Color.Red;
                 MessageBox.Show("Port Bağlantısı Kesildi", "Bağlantı Hatası", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Warning); // Durdur, Yeniden Dene, Yoksay

            }
            catch (Exception)
            {
                // baglanti();
            }  // https://mustafabukulmez.com/2018/01/20/c-messagebox-isinize-yarayacak-tum-ozellikleri/ Burada port bilgisi alınamadığı durumda yapılacak şeyler yazılabilir

        }

        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {/*
            string data;
            string[] splitted_data;
            data = serialPort1.ReadLine();
            splitted_data = data.Split('*');

            label1.Text = splitted_data[0];
            label2.Text = splitted_data[0];
            */
        }
        Task ExcellKayit()
        {
            return Task.Run(() =>
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                excel.Visible = true;

                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);

                Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

                int StartCol = 1;

                int StartRow = 1;

                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];

                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;

                }

                StartRow++;

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {

                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {

                        try
                        {

                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];

                            myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;

                        }

                        catch
                        {



                        }


                    }

                }
            });

            }
        private async void aktar_Click(object sender, EventArgs e)
        {
            // timer1.Enabled = true                                                   // Prograsbar deneme https://www.bilisimkonulari.com/c-progressbar-kullanimi.html
            await ExcellKayit();
            /*
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);

            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            for (int s = 0; s < dataGridView1.Columns.Count; s++)
            {
                sheet1.Cells[1, s + 1] = dataGridView1.Columns[s].HeaderText;
            }

            for (int s = 0; s < dataGridView1.Rows.Count; s++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    sheet1.Cells[s + 2, j + 1] = dataGridView1.Rows[s].Cells[j].Value.ToString();
                }
            }

            /*
            workbook.SaveAs("D:\\ornek.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, Type.Missing, Type.Missing);

            workbook.Close();

            excel.Quit();*/
        }
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void Deger_2_Click(object sender, EventArgs e)
        {

        }

        private void Deger_1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private async void excel_aktar_Click(object sender, EventArgs e)
        {
            await ExcellKayit();
        }

        private void excel_aktar_sil_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void darkmode_CheckedChanged(object sender, EventArgs e)
        {
            if(chk != true)
            {

                chk = true;
                this.BackColor = Color.Black;
                //this.Deger_1.ForeColor = Color.Blue;
                /*
                this.gunaGroupBox7.BaseColor = Color.White;
                this.gunaGroupBox7.BorderColor = Color.White;
                this.gunaGroupBox7.LineColor = Color.White;*/
                for (int gunaGroupBox = 1; gunaGroupBox < 18; gunaGroupBox++)
                {
                    this.ForeColor =Color.White;
                }

            }
            else
            {
                
                chk = false;
                this.BackColor = Color.FromArgb(25, 28, 32);
                //this.Deger_1.ForeColor = Color.Purple;
                //this.gunaGroupBox7.BaseColor = DefaultBackColor;
                // this.gunaGroupBox7.BaseColor = Color.FromArgb(596675);
                /*   this.gunaGroupBox7.BaseColor = Color.DimGray;
                   this.gunaGroupBox7.BorderColor = Color.DimGray;
                   this.gunaGroupBox7.LineColor = Color.DimGray;*/

            }


            /*
            try{
                this.BackColor = Color.Black; //this.ForeColor = Color.Blue;     // FONT RENGİ DEĞİŞMEKTE
            }
            catch (Exception)
            {
                this.BackColor = Color.DarkGray; //this.ForeColor = Color.Blue; // FONT RENGİ DEĞİŞMEKTE
            }
            */
        }

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private void aktar_sil_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            /*
            progressBar1.Increment(1);
            //Timerın başlaması ile birlikte progressbarın increment özelliğini      // Prograsbar deneme
            //kullanıyoruz ve her saniye 1 artıcak şekilde ayarlıyoruz.
            label1.Text = "%" + progressBar1.Value.ToString();
            //Labela progressbarın değerini yazdırıyoruz.
            if (progressBar1.Value == 100)
            {//eğer progressbarın değeri 100e eşitlenirse
                timer1.Stop();//timerı durduruyoruz.
                MessageBox.Show("Yükleme tamamlanmıştır.");
                //Messagebox ile uyarı veriyoruz.
            }*/


            chart1.ChartAreas[0].AxisX.Minimum = minmx;
            chart1.ChartAreas[0].AxisX.Maximum = maksmx;
            chart2.ChartAreas[0].AxisX.Minimum = minmx;
            chart2.ChartAreas[0].AxisX.Maximum = maksmx;
            chart3.ChartAreas[0].AxisX.Minimum = minmx;
            chart3.ChartAreas[0].AxisX.Maximum = maksmx;
            chart4.ChartAreas[0].AxisX.Minimum = minmx;
            chart4.ChartAreas[0].AxisX.Maximum = maksmx;

            chart1.ChartAreas[0].AxisY.Minimum = -120000;
            chart1.ChartAreas[0].AxisY.Maximum = 120000;
            chart2.ChartAreas[0].AxisY.Minimum = -50;
            chart2.ChartAreas[0].AxisY.Maximum = 50;
            chart3.ChartAreas[0].AxisY.Minimum = -20000;
            chart3.ChartAreas[0].AxisY.Maximum = 20000;
            chart4.ChartAreas[0].AxisY.Minimum = -1100;
            chart4.ChartAreas[0].AxisY.Maximum = 1100;

            chart1.ChartAreas[0].AxisX.ScaleView.Zoom(minmx, maksmx);
            chart2.ChartAreas[0].AxisX.ScaleView.Zoom(minmx, maksmx);
            chart3.ChartAreas[0].AxisX.ScaleView.Zoom(minmx, maksmx);
            chart4.ChartAreas[0].AxisX.ScaleView.Zoom(minmx, maksmx);

            serialPort1.Write("1");
            string sonuc = serialPort1.ReadLine();
            string[] pot = sonuc.Split('/');       // split yıldızdan slasha geçirdim.
            //  MessageBox.Show("Port Bağlantısı Kesildi", "Bağlantı Hatası", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Warning);


            satir = dataGridView1.Rows.Add(sonuc);

            dataGridView1.Rows[satir].Cells[0].Value = satirNo;
            dataGridView1.Rows[satir].Cells[7].Value = pot[0];
            dataGridView1.Rows[satir].Cells[8].Value = pot[1];
            dataGridView1.Rows[satir].Cells[9].Value = pot[2];
            dataGridView1.Rows[satir].Cells[10].Value = pot[3];
            dataGridView1.Rows[satir].Cells[11].Value = pot[4];
            dataGridView1.Rows[satir].Cells[12].Value = pot[5];
            dataGridView1.Rows[satir].Cells[6].Value = pot[5];
            dataGridView1.Rows[satir].Cells[7].Value = pot[6];
            dataGridView1.Rows[satir].Cells[8].Value = pot[7];
            dataGridView1.Rows[satir].Cells[9].Value = pot[8];
            dataGridView1.Rows[satir].Cells[10].Value = pot[9];
            dataGridView1.Rows[satir].Cells[11].Value = pot[10];
            dataGridView1.Rows[satir].Cells[12].Value = pot[11];                            // enlem boylam verisi eksik sadece textbox'a eklenecek

            dataGridView1.Rows[satir].Cells[13].Value = DateTime.Now.ToLongTimeString();    // EN BAŞA STRİNG OLARAK EKLEDİĞİMİZ KOD KISMINDA SÜRE İLERLEMEDİĞİ İÇİN UZUN ZAMAN ZARFI OLARAK DATATİME I BURAYA EKLEDİM
            dataGridView1.Rows[satir].Cells[14].Value = yeni.ToShortDateString();
            satir++;
            satirNo++;


            if (sonuc != null)
            {

                /*
                sicaklik = pot[0];                                                              // BU ŞEKİLDE OLACAK
                textBox1.Text = pot[0];
                */
          

                   // textBox1.Text = sonuc + "";
                   textBox7.Text = pot[0] + "";
                   textBox8.Text = pot[1] + "";
                   textBox9.Text = pot[2] + "";
                   textBox10.Text = pot[3] + "";
                   textBox11.Text = pot[4] + "";
                    textBox12.Text = pot[5] + "";
                    textBox6.Text = pot[5] + "";
                    textBox7.Text = pot[6] + "";
                    textBox8.Text = pot[7] + "";
                    textBox9.Text = pot[8] + "";
                    textBox10.Text = pot[9] + "";
                    textBox11.Text = pot[10] + "";
                    textBox12.Text = pot[11] + "";
                    // enlem boylam verisi eksik sadece textbox'a eklenecek

                 this.chart1.Series[0].Points.AddXY((minmx + maksmx) / 2, pot[1]);
                 this.chart2.Series[0].Points.AddXY((minmx + maksmx) / 2, pot[0]);
                 this.chart3.Series[0].Points.AddXY((minmx + maksmx) / 2, pot[10]);             // sonra eklenecek 2 uçuş kodu birleştiğindepu ile chart edilecek
                 this.chart4.Series[0].Points.AddXY((minmx + maksmx) / 2, pot[2]);
                maksmx++;
                minmx++;

            }

            serialPort1.DiscardInBuffer();

        }
    }

    
}
