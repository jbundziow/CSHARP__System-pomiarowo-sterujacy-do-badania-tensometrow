/** 
  * Plik            :   Form1.cs
  * Data ost. mod.  :   22.10.2021, 18:30
  * 
  * Autor           :   Jakub Bundziów
  * Data wykonania  :   2021
  * Kierunek        :   Automatyka i Robotyka
  * Specjalizacja   :   Automatyka Przemysłowa
  * Cel             :   Aplikacja stworzona na potrzeby pracy inżynierskiej pt. "System pomiarowo-sterujący do badania tensometrów"
  * Promotor        :   dr inż. Leszek Furmankiewicz
  *
  **/

/**
 * UKŁAD STEROWANIA ARDUINO - POLECENIA
 * 
 * 0 - napięcia odniesienia OFF (pomiar rezystancji)
 * 1 - napięcie odniesienia 2V
 * 2 - napięcie odniesienia 4V
 * 3 - odczyt z czujnika przemieszczenia
 * 4 - polecenie ruchu silnika do góry - 1 impuls
 * 5 - polecenie ruchu silnika w dół   - 1 impuls
 * 6 - silniki OFF
 * 
 **/

//BIBLIOTEKI
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using Ivi.Visa.Interop;


namespace SystemPomiarowoSterujacyTensometry
{
    public partial class Form1 : Form
    {

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      ZMIENNE GLOBALNE
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        //komunikacja z urządzeniami
        SerialPort sp_ukl_ster;
        SerialPort sp_multimetr;


        //status połączenia z urządzeniami
        public bool status_ukl_ster;
        public bool status_multimetr;


        //zmienne aktywowane w trakcie ręcznego poruszania silnikiem góra/dół
        public bool silnik_gora_klik;
        public bool silnik_dol_klik;


        //aktualnie ustawione wartości w comboBox'ach
        public int stan_comboBox_zasilanie_mostka;
        public int stan_comboBox_rodzaj_pomiaru;


        //zmienne do obsługi funkcjonalności programu
        public int lp = 1; //liczba porządkowa do numerowania zbieranych pomiarów
        public double skalibrowany_czuj_przem = 999; //wartość napięcia na wyjściu czujnika przy ugięciu 0mm (kalibracja przez użytkownika i zapis w tej zmiennej)
        public double ustawione_ugiecie; //ugięcie ustawione przez użytkownika w oknie "Pomiary"
        public bool polecenie_ugnij_belke;
        public bool czy_zapisano_ustawienia;
        public bool czy_zerwano_polaczenie_w_trakcie_wykonywania_pomiarow;
        public bool co_drugi_pomiar; //w celu nie zawieszania się dokonywania odczytów odległości przy poruszaniu silnikiem
        public bool stop_motor; //zmienna do awaryjnego zatrzymania silnika


        //zmienne do MessageBox'ów z wyborem YES/NO
        DialogResult zmiana_ust_czy_kasujemy_postep; //DialogResult.Yes lub DialogResult.No
        DialogResult czyszczenie_wykresu_i_tabeli_czy_kasujemy_postep; //DialogResult.Yes lub DialogResult.No
        DialogResult czy_chcesz_zapisac_pomiar; //DialogResult.Yes lub DialogResult.No
        DialogResult czy_odlaczono_zasilanie; //DialogResult.Yes lub DialogResult.No


        //zmienne do multimetru Keysight
        //Form1 DmmClass = new Form1(); //Create an instance of this class so we can call functions from Main
        Ivi.Visa.Interop.ResourceManager rm = new Ivi.Visa.Interop.ResourceManager(); //Open up a new resource manager
        Ivi.Visa.Interop.FormattedIO488 myDmm = new Ivi.Visa.Interop.FormattedIO488(); //Open a new Formatted IO 488 session 


        //tytuł wykresu - obiekt
        Title title;


        //zmienne dzięki którym, okno aplikacji może być przesuwane przez użytkownika po ekranie
        int mov;
        int movX;
        int movY;

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------










        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      WYGLĄD ELEMENTÓW
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        //OBRAMOWANIE WOKÓŁ PANELI
        private void panel_multimetr_Paint(object sender, PaintEventArgs e) //Okno "Komunikacja" -> połączenie z multimetrem
        {
            ControlPaint.DrawBorder(e.Graphics, panel_multimetr.ClientRectangle,
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // left
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // top
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // right
            Color.SkyBlue, 3, ButtonBorderStyle.Solid);// bottom
        }

        void panel_ukl_ster_Paint(object sender, PaintEventArgs e) //Okno "Komunikacja" -> połączenie z układem sterowania
        {
            ControlPaint.DrawBorder(e.Graphics, panel_ukl_ster.ClientRectangle,
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // left
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // top
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // right
            Color.SkyBlue, 3, ButtonBorderStyle.Solid);// bottom
        }

        private void panel_testowanie_multimetr_Paint(object sender, PaintEventArgs e) //Okno "Testowanie" -> testowanie multimetru
        {
            ControlPaint.DrawBorder(e.Graphics, panel_testowanie_multimetr.ClientRectangle,
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // left
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // top
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // right
            Color.SkyBlue, 3, ButtonBorderStyle.Solid);// bottom
        }

        private void panel_testowanie_ukl_ster_Paint(object sender, PaintEventArgs e) //Okno "Testowanie" -> testowanie układu sterowania
        {
            ControlPaint.DrawBorder(e.Graphics, panel_testowanie_ukl_ster.ClientRectangle,
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // left
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // top
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // right
            Color.SkyBlue, 3, ButtonBorderStyle.Solid);// bottom
        }

        private void panel_pomiary_ustawienia_Paint(object sender, PaintEventArgs e) //Okno "Pomiary" -> panel z ustawieniami początkowymi
        {
            ControlPaint.DrawBorder(e.Graphics, panel_pomiary_ustawienia.ClientRectangle,
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // left
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // top
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // right
            Color.SkyBlue, 3, ButtonBorderStyle.Solid);// bottom
        }

        private void panel_pomiary_pomiary_Paint(object sender, PaintEventArgs e) //Okno "Pomiary" -> panel z elementami związanymi z przeprowadzaniem pomiarów
        {
            ControlPaint.DrawBorder(e.Graphics, panel_pomiary_pomiary.ClientRectangle,
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // left
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // top
            Color.SkyBlue, 3, ButtonBorderStyle.Solid, // right
            Color.SkyBlue, 3, ButtonBorderStyle.Solid);// bottom
        }

        //EFEKT NACIŚNIĘCIA PRZYCISKÓW ZWIĄZANYCH Z RUCHEM SILNIKA GÓRA/DÓŁ
        private void pictureBox_silnik_gora_pom_MouseDown(object sender, MouseEventArgs e) //BUTTON POMIARY SILNIK GÓRA MouseDown
        {
            pictureBox_silnik_gora_pom.Image = Properties.Resources.scroll_up_clicked_32px; //zmień obrazek na clicked
        }

        private void pictureBox_silnik_gora_pom_MouseUp(object sender, MouseEventArgs e) //BUTTON POMIARY SILNIK GÓRA MouseUp
        {
            pictureBox_silnik_gora_pom.Image = Properties.Resources.scroll_up_32px; //zmień obrazek na normalny
        }

        private void pictureBox_silnik_dol_pom_MouseDown(object sender, MouseEventArgs e) //BUTTON POMIARY SILNIK DÓŁ MouseDown
        {
            pictureBox_silnik_dol_pom.Image = Properties.Resources.scroll_down_clicked_32px; //zmień obrazek na clicked
        }

        private void pictureBox_silnik_dol_pom_MouseUp(object sender, MouseEventArgs e) //BUTTON POMIARY SILNIK DÓŁ MouseUp
        {
            pictureBox_silnik_dol_pom.Image = Properties.Resources.scroll_down_32px; //zmień obrazek na normalny
        }

        private void pictureBox_silnik_gora_MouseDown(object sender, MouseEventArgs e) //BUTTON TESTOWANIE SILNIK GÓRA MouseDown
        {
            pictureBox_silnik_gora.Image = Properties.Resources.scroll_up_clicked_32px; //zmień obrazek na clicked
        }

        private void pictureBox_silnik_gora_MouseUp(object sender, MouseEventArgs e) //BUTTON TESTOWANIE SILNIK GÓRA MouseUp
        {
            pictureBox_silnik_gora.Image = Properties.Resources.scroll_up_32px; //zmień obrazek na normalny
        }

        private void pictureBox_silnik_dol_MouseDown(object sender, MouseEventArgs e) //BUTTON TESTOWANIE SILNIK DÓŁ MouseDown
        {
            pictureBox_silnik_dol.Image = Properties.Resources.scroll_down_clicked_32px; //zmień obrazek na clicked
        }

        private void pictureBox_silnik_dol_MouseUp(object sender, MouseEventArgs e) //BUTTON TESTOWANIE SILNIK DÓŁ MouseUp
        {
            pictureBox_silnik_dol.Image = Properties.Resources.scroll_down_32px; //zmień obrazek na normalny
        }

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------









        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      PODSTAWOWE ELEMENTY KONSTRUKCJI APLIKACJI
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        public Form1() //ZAŁADOWANIE OKNA APLIKACJI FORM1
        {
            InitializeComponent();
            label_data_i_czas.Text = System.DateTime.Now.ToString();


            //USTAWIENIA POCZĄTKOWE DLA TAB CONTROL
            tabControl.SendToBack();
            tabControl.Location = new System.Drawing.Point(175, 100);
            tabControl.Size = new System.Drawing.Size(1250, 680); //schowanie bocznego paska


            //USTAWIENIA - niebieski panel, który wizualizuje obecność w danej zakładce
            panel_akt_okno.Height = button_mainpage.Height;
            panel_akt_okno.Top = button_mainpage.Top;
            //wywołanie strony głównej na starcie
            tabControl.SelectTab("tabPage1"); //strona główna


            comboBox_zasilanie_mostka.SelectedIndex = 0; //domyślna wartość zasilania mostka OFF
            comboBox_rodzaj_pomiaru.SelectedIndex = 0; //domyślna wartość 'pomiar rezystancji'


            //USTAWIENIA POCZĄTKOWE DLA WYKRESU
            chart1.Series.Clear();
            chart1.Series.Add("Multimetr");
            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line; //typ wykresu - liniowy
            chart1.Series[0].IsVisibleInLegend = false; //brak legendy

            //USTAWIENIA POCZĄTKOWE DLA TEXTBOX'OW W OKNIE "INFORMACJE"
            textBox1.Cursor = Cursors.Arrow;
            textBox1.GotFocus += textBox_GotFocus;
            textBox2.Cursor = Cursors.Arrow;
            textBox2.GotFocus += textBox_GotFocus;
            textBox3.Cursor = Cursors.Arrow;
            textBox3.GotFocus += textBox_GotFocus;

            //znaczniki dla punktów - ustawienia wyglądu
            chart1.Series[0].MarkerBorderColor = Color.Black;
            chart1.Series[0].MarkerColor = Color.Gray;
            chart1.Series[0].MarkerSize = 6;
            chart1.Series[0].MarkerStyle = MarkerStyle.Circle;

            //tytuł wykresu, podpisanie osi
            chart1.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Calibri Light", 14.00F, System.Drawing.FontStyle.Regular);
            chart1.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Calibri Light", 14.00F, System.Drawing.FontStyle.Regular);
            chart1.Titles.Clear();
            title = chart1.Titles.Add("Charakterystyka rezystancji tensometru w funkcji ugięcia belki lub\ncharakterystyka napięcia w przekątnej mostka w funkcji ugięcia belki");
            title.Font = new System.Drawing.Font("Calibri Light", 12, FontStyle.Regular);
            title.Alignment = ContentAlignment.MiddleCenter;
            chart1.ChartAreas[0].AxisX.Title = "Ugięcie belki [mm]";
            chart1.ChartAreas[0].AxisY.Title = "R [Ω] lub U [mV]";
            chart1.ChartAreas[0].AxisY.IsStartedFromZero = false;
            chart1.ChartAreas[0].AxisX.IsStartedFromZero = false;

            chart1.Series["Multimetr"].Points.AddXY(0, 0); //punkt na początek (0,0)


            //USTAWIENIA DLA DATAGRIDVIEW (TABLEA)
            this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 15);
            this.dataGridView1.DefaultCellStyle.ForeColor = Color.Black;
            this.dataGridView1.DefaultCellStyle.BackColor = Color.Beige;
            this.dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Yellow;
            this.dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Black;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;


            listbox_historia_add("Uruchomienie aplikacji."); //wysłanie komunikatu

           
        }


        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        // ********* PRZESUWANIE OKNA, MINIMALIZOWANIE APLIKACJI, ZAMYKANIE APLIKACJI **********
        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        private void button_zamknij_app_Click(object sender, EventArgs e) //ZAMYKANIE APLIKACJI
        {
            reset_ukl_ster();
            System.Windows.Forms.Application.Exit(); //zamknij aplikację
        }

        private void button_minimalizuj_app_Click(object sender, EventArgs e) //MINIMALIZACJA APLIKACJI
        {
            this.WindowState = FormWindowState.Minimized; //zminimalizuj aplikację
        }



        //PRZESUWANIE OKNA PO EKRANIE
        private void panel_gorny_pasek_MouseDown(object sender, MouseEventArgs e) //PRZESUWANIE - funkcja 1/3
        {
            mov = 1;
            movX = e.X;
            movY = e.Y;
        }

        private void panel_gorny_pasek_MouseMove(object sender, MouseEventArgs e) //PRZESUWANIE - funkcja 2/3
        {
            if (mov == 1)
            {
                this.SetDesktopLocation(MousePosition.X - movX, MousePosition.Y - movY);
            }
        }

        private void panel_gorny_pasek_MouseUp(object sender, MouseEventArgs e) //PRZESUWANIE - funkcja 3/3
        {
            mov = 0;
        }





        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        // ********* MENU **********
        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 

        private void button_mainpage_Click(object sender, EventArgs e) //PRZYCISK "STRONA GŁÓWNA"
        {
            //USTAWIENIA - niebieski panel, który wizualizuje obecność w danej zakładce
            panel_akt_okno.Height = button_mainpage.Height;
            panel_akt_okno.Top = button_mainpage.Top;
            tabControl.SelectTab("tabPage1"); //strona główna

            reset_ukl_ster();
            timer_test.Enabled = false;
            timer_pomiary_ustawienia.Enabled = false;
            timer_wspolrzedne.Enabled = false;
        }

        private void button_komunikacja_Click(object sender, EventArgs e) //PRZYCISK "KOMUNIKACJA"
        {
            panel_akt_okno.Height = button_komunikacja.Height;
            panel_akt_okno.Top = button_komunikacja.Top;
            tabControl.SelectTab("tabPage2"); //komunikacja

            reset_ukl_ster();
            timer_test.Enabled = false;
            timer_pomiary_ustawienia.Enabled = false;
            timer_wspolrzedne.Enabled = false;
        }

        private void button_testowanie_Click(object sender, EventArgs e) //PRZYCISK "TESTOWANIE"
        {
            panel_akt_okno.Height = button_testowanie.Height;
            panel_akt_okno.Top = button_testowanie.Top;
            tabControl.SelectTab("tabPage3"); //testowanie

            timer_wspolrzedne.Enabled = false;

            comboBox_zasilanie_mostka.SelectedIndex = 0; //domyślna wartość OFF
            reset_ukl_ster();
            if (czy_zapisano_ustawienia == true) //jeśli zapisano ustawienia == pomiary aktualnie trwają
                timer_pomiary_pomiary.Enabled = false; //wyłącz pomiary na czas bycia w zakładce "testowanie"

            timer_pomiary_ustawienia.Enabled = false;

            if (status_ukl_ster == true) //jeśli połączono z układem sterowania
                timer_test.Enabled = true; //włącz timer odpowiedzialny za wysyłanie danych do układu sterowania
            else
                timer_test.Enabled = false; //wyłącz timer odpowiedzialny za wysyłanie danych do układu sterowania

        }

        private void button_pomiary_Click(object sender, EventArgs e) //PRZYCISK "POMIARY"
        {
            panel_akt_okno.Height = button_pomiary.Height;
            panel_akt_okno.Top = button_pomiary.Top;
            tabControl.SelectTab("tabPage4"); //komunikacja

            reset_ukl_ster();
            timer_test.Enabled = false;
            timer_wspolrzedne.Enabled = true; //możliwość zaznaczania współrzędnych punktów na wykresie włączona


            //jeśli już zapisano wcześniej ustawienia, to umożliw dalsze wykonywanie pomiarów po powrocie do zakładki "Pomiary"
            if (czy_zapisano_ustawienia == true)
            {
                timer_pomiary_pomiary.Enabled = true;
            }


            if (status_ukl_ster == true)
            {
                panel_pomiary_ustawienia.Enabled = true; //aktywacja panela "pomiary-ustawinia"
                if (skalibrowany_czuj_przem == 999) //jeśli nieskalibrowano jeszcze czujnika przemieszczenia w zakładce 'pomiary'
                    timer_pomiary_ustawienia.Enabled = true; //włącz timer
            }
            else //jeśli niepołączono z układem sterowania uniemożliw wykonywanie żadnych akcji
            {
                panel_pomiary_ustawienia.Enabled = false;
                timer_pomiary_ustawienia.Enabled = false;
            }
        }

        private void button_historia_Click(object sender, EventArgs e) //PRZYCISK "HISTORIA"
        {
            panel_akt_okno.Height = button_historia.Height;
            panel_akt_okno.Top = button_historia.Top;
            tabControl.SelectTab("tabPage5"); //komunikacja

            reset_ukl_ster();
            timer_test.Enabled = false;
            timer_pomiary_ustawienia.Enabled = false;
            timer_wspolrzedne.Enabled = false;
        }

        private void button_informacje_Click(object sender, EventArgs e) //PRZYCISK "INFORMACJE"
        {
            panel_akt_okno.Height = button_informacje.Height;
            panel_akt_okno.Top = button_informacje.Top;
            tabControl.SelectTab("tabPage6"); //komunikacja

            reset_ukl_ster();
            timer_test.Enabled = false;
            timer_pomiary_ustawienia.Enabled = false;
            timer_wspolrzedne.Enabled = false;
        }

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------










        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      OBSŁUGA ZAKŁADKI "STRONA GŁÓWNA"
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        //PUSTO

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------










        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      OBSŁUGA ZAKŁADKI "KOMUNIKACJA"
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        // ********* UKŁAD STEROWANIA **********
        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        private void button_ukl_ster_polacz_Click(object sender, EventArgs e) //UKŁAD STEROWANIA - PRZYCISK POŁĄCZ
        {
            //PARAMETRY KOMUNIKACJI - UKŁAD STEROWANIA ARDUINO
            sp_ukl_ster = new SerialPort();
            sp_ukl_ster.PortName = textBox_port_com_ukl_ster.Text; //pobranie portu COM z TextBox'a
            sp_ukl_ster.BaudRate = 9600;
            sp_ukl_ster.Parity = Parity.None;
            sp_ukl_ster.DataBits = 8;
            sp_ukl_ster.StopBits = StopBits.One;
            sp_ukl_ster.Handshake = Handshake.None;
            sp_ukl_ster.ReadTimeout = 500; //wprowadzenie timeoutu odczytu
            sp_ukl_ster.WriteTimeout = 500; //wprowadzenie timeoutu wysyłania danych

            try
            {
                sp_ukl_ster.Open(); //otwarcie portu
                listbox_historia_add("Połączono z układem sterowania. Port: " + textBox_port_com_ukl_ster.Text + ".");
                
                //obrazki OK
                pictureBox_ukl_ster_status1.Image = Properties.Resources.ok_25px;
                pictureBox_ukl_ster_status2.Image = Properties.Resources.ok_25px;

                button_ukl_ster_polacz.Enabled = false;
                label_ukl_ster_error.Visible = false;
                panel_testowanie_ukl_ster.Enabled = true;
                status_ukl_ster = true;
                textBox_port_com_ukl_ster.ReadOnly = true;

                if (czy_zerwano_polaczenie_w_trakcie_wykonywania_pomiarow == true)
                {
                    timer_pomiary_pomiary.Enabled = true;
                    czy_zerwano_polaczenie_w_trakcie_wykonywania_pomiarow = false;
                }

            }
            catch
            {
                button_ukl_ster_rozlacz_Click(this, EventArgs.Empty);
                listbox_historia_add("Błąd podczas łączenia z układem sterowania! Czy wpisany port '" + textBox_port_com_ukl_ster.Text + "' jest prawidłowy?");

                //error pod panelem
                label_ukl_ster_error.Text = "Błąd podczas łączenia z układem sterowania!\nCzy wpisany port '" + textBox_port_com_ukl_ster.Text + "' jest prawidłowy?";
                label_ukl_ster_error.Visible = true;
            }
            reset_ukl_ster();
        }

        private void button_ukl_ster_rozlacz_Click(object sender, EventArgs e) //UKŁAD STEROWANIA - PRZYCISK ROZŁĄCZ
        {
            sp_ukl_ster.Close(); //zamknięcie portu
            listbox_historia_add("Rozłączono z układem sterowania.");

            //obrazki rozłączono
            pictureBox_ukl_ster_status1.Image = Properties.Resources.cancel_25px;
            pictureBox_ukl_ster_status2.Image = Properties.Resources.cancel_25px;

            button_ukl_ster_polacz.Enabled = true;
            label_ukl_ster_error.Visible = false;
            panel_testowanie_ukl_ster.Enabled = false;
            status_ukl_ster = false;
            textBox_port_com_ukl_ster.ReadOnly = false;
        }

        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        // ********* MULTIMETR **********
        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 

        private void button_keysight_polacz_Click(object sender, EventArgs e) //MULTIMETR - PRZYCISK POŁĄCZ
        {


            try
            {
                string DutAddr = "USB0::0x0957::0x0607::MY47000832::0::INSTR"; //String for USB

                myDmm.IO = (IMessage)rm.Open(DutAddr, AccessMode.NO_LOCK, 2000, ""); //Open up a handle to the DMM with a 2 second timeout
                myDmm.IO.Timeout = 3000; //You can also set your timeout by doing this command, sets to 3 seconds
                //First start off with a reset state
                myDmm.IO.Clear(); //Send a device clear first to stop any measurements in process
                myDmm.WriteString("*RST", true); //Reset the device


                listbox_historia_add("Połączono z multimetrem. Adres: " + textBox_port_com_multimetr.Text + ".");

                //obrazki OK
                pictureBox_multimetr_status1.Image = Properties.Resources.ok_25px;
                pictureBox_multimetr_status2.Image = Properties.Resources.ok_25px;

                button_keysight_polacz.Enabled = false;
                label_multimetr_error.Visible = false;
                panel_testowanie_multimetr.Enabled = true;
                status_multimetr = true;
                textBox_port_com_multimetr.ReadOnly = true;

                if (czy_zerwano_polaczenie_w_trakcie_wykonywania_pomiarow == true)
                {
                    timer_pomiary_pomiary.Enabled = true;
                    czy_zerwano_polaczenie_w_trakcie_wykonywania_pomiarow = false;
                }

            }
            catch (Exception exception)
            {
                button_keysight_rozlacz_Click(this, EventArgs.Empty);
                listbox_historia_add("Błąd podczas łączenia z multimetrem! Kod błędu: " +exception.Message);

                //error pod panelem
                label_multimetr_error.Text = "Błąd podczas łączenia z multimetrem!";
                label_multimetr_error.Visible = true;


            }
        }

        private void button_keysight_rozlacz_Click(object sender, EventArgs e) //MULTIMETR - PRZYCISK ROZŁĄCZ
        {
            //Close out your resources
            //try { myDmm.IO.Clear(); }
            //catch { }
            //try { myDmm.IO.Close(); }
            //catch { }
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(myDmm); }
            catch { }
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rm);
            }
            catch { }

            listbox_historia_add("Rozłączono z multimetrem.");
            autoscroll_listbox("listBox_historia");


            //obrazki rozłączono
            pictureBox_multimetr_status1.Image = Properties.Resources.cancel_25px;
            pictureBox_multimetr_status2.Image = Properties.Resources.cancel_25px;

            button_keysight_polacz.Enabled = true;
            label_multimetr_error.Visible = false;
            panel_testowanie_multimetr.Enabled = false;
            status_multimetr = false;
            textBox_port_com_multimetr.ReadOnly = false;
        }

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------










        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      OBSŁUGA ZAKŁADKI "TESTOWANIE"
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        // ********* UKŁAD STEROWANIA **********
        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        private void pictureBox_silnik_gora_Click(object sender, EventArgs e) //kliknięcie przycisku "silnik góra"
        {
            silnik_gora_klik = true;
            listbox_historia_add("Wysłano do układu sterowania polecenie ręcznego ruchu silnika w górę.");
        }

        private void pictureBox_silnik_dol_Click(object sender, EventArgs e) //kliknięcie przycisku "silnik dół"
        {
            silnik_dol_klik = true;
            listbox_historia_add("Wysłano do układu sterowania polecenie ręcznego ruchu silnika w dół.");
        }

        private void comboBox_zasilanie_mostka_SelectedIndexChanged(object sender, EventArgs e) //zmiana w comboBox - "Zasilanie mostka"
        {
            stan_comboBox_zasilanie_mostka = comboBox_zasilanie_mostka.SelectedIndex; //przypisanie wybranego indeksu do zmiennej
            reset_ukl_ster();
        }


        private void timer_test_Tick(object sender, EventArgs e) //TIMER DO KOMUNIKACJI Z URZ. POM. W TRAKCIE TESTOWANIA
        {
            try
            {
                switch (stan_comboBox_zasilanie_mostka)
                {

                    case 0: //4V
                        sp_ukl_ster.WriteLine("1");

                        button_multimetr_pomiar_napiecia.Enabled = true;
                        button_multimetr_pomiar_rezystancji.Enabled = true;
                        break;

                    case 1: //2V
                        sp_ukl_ster.WriteLine("2");

                        button_multimetr_pomiar_napiecia.Enabled = true;
                        button_multimetr_pomiar_rezystancji.Enabled = true;
                        break;

                    

                    default: //OFF
                        sp_ukl_ster.WriteLine("0");

                        button_multimetr_pomiar_napiecia.Enabled = true;
                        button_multimetr_pomiar_rezystancji.Enabled = true;
                        break;
                }


                //150
                int ilosc_impulsow = 150; //@@@ do ustawienia długość ruchu w trakcie testowania - ruch góra/dół
                if (silnik_gora_klik == true)
                {
                    for (int n = 0; n < ilosc_impulsow; n++)
                    {
                        sp_ukl_ster.WriteLine("4");
                    }

                    silnik_gora_klik = false;
                }

                if (silnik_dol_klik == true)
                {
                    for (int n = 0; n < ilosc_impulsow; n++)
                    {
                        sp_ukl_ster.WriteLine("5");
                    }
                    silnik_dol_klik = false;
                }


                //odczytanie napięcia z czujnika przemieszczenia
                //sp_ukl_ster.DiscardInBuffer(); //czyszczenie bufora (nie działa)
                sp_ukl_ster.WriteLine("3");
                //textBox_odczyt_czuj_przem_V_test.Text = sp_ukl_ster.ReadExisting(); //stara metoda - zapychał się bufor, pojawiały się śmieci
                string odczyt = sp_ukl_ster.ReadExisting();
                
                try
                {
                    odczyt = (double.Parse(odczyt, System.Globalization.CultureInfo.InvariantCulture)*5/1024).ToString(); //przeliczenie liczby od 0-1024 na napięcie
                
                
                if (!String.IsNullOrWhiteSpace(odczyt) && (odczyt.Length >= 4) && (odczyt[1] == ',')) //jeśli odczyt nie jest pusty/ ma minimum 4 znaki, a na drugiem miejscu jest przecinek
                {
                    textBox_odczyt_czuj_przem_V_test.Text = odczyt.Substring(0, 4); //umieść w textBox pierwsze 4 znaki z bufora, np. 3.29
                }
                }
                catch { }  //nie wyrzucaj bledu jesli jeden z pomiarow sie nie powiedzie

            } //koniec try

            catch
            {
                button_ukl_ster_rozlacz_Click(this, EventArgs.Empty);
                timer_test.Enabled = false;
                listbox_historia_add("Połączenie z układem sterowania zostało zerwane.");
                MessageBox.Show("Połączenie z układem sterowania zostało zerwane.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        // ********* MULTIMETR **********
        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 

        private void button_multimetr_identyfikacja_przyrzadu_Click(object sender, EventArgs e)
        {
            try
            {
                
                listbox_historia_add("Wysłano do multimetru polecenie identyfikacji przyrządu.");
                myDmm.IO.Clear(); //Send a device clear first to stop any measurements in process
                myDmm.WriteString("*RST", true); //Reset the device
                myDmm.WriteString("*IDN?", true); //Get the IDN string                
                string IDN = myDmm.ReadString();
                listbox_odp_multimetr_add(IDN);
                listbox_historia_add("TESTOWANIE/Multimetr - identyfikacja przyrządu: " + IDN);

            }
            catch
            {
                button_keysight_rozlacz_Click(this, EventArgs.Empty);
                listbox_historia_add("Połączenie z multimetrem zostało zerwane.");
                MessageBox.Show("Połączenie z multimetrem zostało zerwane.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button_multimetr_pomiar_napiecia_Click(object sender, EventArgs e)
        {
            try
            {
                listbox_historia_add("Wysłano do multimetru polecenie pomiaru napięcia.");
                //Configure for DCV 100V range, 100uV resolution
                //myDmm.WriteString("CONF:VOLT:DC 100, 0.0001", true);

                //Configure for DCV 0.2V range, 10uV resolution
                myDmm.WriteString("*RST", true); //Reset the device
                //myDmm.WriteString("CONF:mVOLT:DC 10, 0.001", true);
                myDmm.WriteString("CONF:VOLT:DC", true);
                myDmm.WriteString("READ?", true);
                string DCVResult = myDmm.ReadString();
                double napiecie_odczyt = Convert.ToDouble(DCVResult, System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
                listbox_odp_multimetr_add("Odczytane napięcie = " + napiecie_odczyt +" [V]"); //report the DCV reading
                listbox_historia_add("TESTOWANIE/Multimetr - odczyt napięcia: " + napiecie_odczyt +" [V]");

            }
            catch
            {
                button_keysight_rozlacz_Click(this, EventArgs.Empty);
                listbox_historia_add("Połączenie z multimetrem zostało zerwane.");
                MessageBox.Show("Połączenie z multimetrem zostało zerwane.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button_multimetr_pomiar_rezystancji_Click(object sender, EventArgs e)
        {
            try
            {
                listbox_historia_add("Wysłano do multimetru polecenie pomiaru rezystancji.");
                //Configure for OHM 2 wire 100 Ohm range, 100uOhm resolution
                //myDmm.WriteString("CONF:RES 100, 0.0001", true);

                //Configure for OHM 2 wire 300 Ohm range, 1mOhm resolution
                myDmm.WriteString("*RST", true); //Reset the device
                myDmm.WriteString("CONF:RES 300, 0.0001", true);
                myDmm.WriteString("READ?", true);
                string Res2WResult = myDmm.ReadString();
                double rezystancja_odczyt = Convert.ToDouble(Res2WResult, System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
                listbox_odp_multimetr_add("Odczytana rezystancja = " + rezystancja_odczyt + " [Ω]"); //report the DCV reading
                listbox_historia_add("TESTOWANIE/Multimetr - odczyt rezystancji: " + rezystancja_odczyt + " [Ω]");

            }
            catch
            {
                button_keysight_rozlacz_Click(this, EventArgs.Empty);
                listbox_historia_add("Połączenie z multimetrem zostało zerwane.");
                MessageBox.Show("Połączenie z multimetrem zostało zerwane.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button_listBox_odp_multimetr_clear_Click(object sender, EventArgs e) //CZYSZCZENIE LISTY ODCZYTÓW Z MULTIMETRU
        {
            listBox_odp_multimetr.Items.Clear();
        }


        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------










        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      OBSŁUGA ZAKŁADKI "POMIARY"
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        // ********* CZĘŚĆ 1 - USTAWIENIA **********
        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 

        private void comboBox_rodzaj_pomiaru_SelectedIndexChanged(object sender, EventArgs e) //zmiana w comboBox - "Rodzaj pomiaru"
        {
            stan_comboBox_rodzaj_pomiaru = comboBox_rodzaj_pomiaru.SelectedIndex; //przypisanie wybranego indeksu do zmiennej
            reset_ukl_ster();
        }

        private void pictureBox_silnik_gora_pom_Click(object sender, EventArgs e) //kliknięcie przycisku "silnik góra"
        {
            silnik_gora_klik = true;
            listbox_historia_add("Wysłano do układu sterowania polecenie ręcznego ruchu silnika w górę.");
        }

        private void pictureBox_silnik_dol_pom_Click(object sender, EventArgs e) //kliknięcie przycisku "silnik dół"
        {
            silnik_dol_klik = true;
            listbox_historia_add("Wysłano do układu sterowania polecenie ręcznego ruchu silnika w dół.");
        }

        private void timer_pomiary_ustawienia_Tick(object sender, EventArgs e) //TIMER DO KOMUNIKACJI Z UKŁADEM STEROWANIA PODCZAS WPROWADZANIA USTAWIEŃ
        {
            try
            {
                sp_ukl_ster.WriteLine("3"); //odczyt z czujnika przemieszczenia
                //textBox_odczyt_czuj_przem_V.Text = sp_ukl_ster.ReadExisting(); //stara metoda - zapychał się bufor

                string odczyt = sp_ukl_ster.ReadExisting();

                try
                {
                    odczyt = (double.Parse(odczyt, System.Globalization.CultureInfo.InvariantCulture) * 5 / 1024).ToString(); //przeliczenie liczby od 0-1024 na napięcie



                if (!String.IsNullOrWhiteSpace(odczyt) && (odczyt.Length >= 4) && (odczyt[1] == ',')) //jeśli odczyt nie jest pusty/ ma minimum 4 znaki, a na drugiem miejscu jest przecinek
                {
                    textBox_odczyt_czuj_przem_V.Text = odczyt.Substring(0, 4); //umieść w textBox pierwsze 4 znaki z bufora, np. 3.29
                }

            }
            catch { }  //nie wyrzucaj bledu jesli jeden z pomiarow sie nie powiedzie

                if (silnik_gora_klik == true)
                {

                    for (int n = 0; n < numericUpDown_dlugosc_ruchu.Value*30; n++) //@@@ w razie potrzeby dostosować długość ruchu 
                    {
                        sp_ukl_ster.WriteLine("4");
                    }

                    silnik_gora_klik = false;

                }

                if (silnik_dol_klik == true)
                {
                    for (int n = 0; n < numericUpDown_dlugosc_ruchu.Value*30; n++)
                    {
                        sp_ukl_ster.WriteLine("5");
                    }
                    silnik_dol_klik = false;

                }

            } //koniec try

            catch
            {
                timer_pomiary_ustawienia.Enabled = false;
                panel_pomiary_ustawienia.Enabled = false;
                listbox_historia_add("Połączenie z układem sterowania zostało zerwane w trakcie zmieniania ustawień w zakładce 'Pomiary'.");
                MessageBox.Show("Połączenie z układem sterowania zostało zerwane.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button_ukl_ster_rozlacz_Click(this, EventArgs.Empty);
            }
        }

        private void button_kalibruj_Click(object sender, EventArgs e)
        {
            try
            {
                //skalibrowany_czuj_przem = double.Parse(textBox_odczyt_czuj_przem_V.Text, System.Globalization.CultureInfo.InvariantCulture); //zapisanie do zmiennej
                skalibrowany_czuj_przem = double.Parse(textBox_odczyt_czuj_przem_V.Text); //zapisanie do zmiennej

                timer_pomiary_ustawienia.Enabled = false; //zatrzymanie dalszych odczytów
                button_kalibruj.Enabled = false;
                pictureBox_silnik_gora_pom.Enabled = false;
                pictureBox_silnik_dol_pom.Enabled = false;
                numericUpDown_dlugosc_ruchu.Enabled = false;
                textBox_odczyt_czuj_przem_V.BackColor = Color.LightGreen; //kolor zielony (potwierdzenie skalibrowania)
                textBox_odczyt_czuj_przem_V.TextAlign = HorizontalAlignment.Center;

                listbox_historia_add("Skalibrowano czujnik przemieszczenia. Efekt:    0 mm = " + skalibrowany_czuj_przem + " V.");
            }
            catch
            {
                listbox_historia_add("Kalibracja czujnika przemieszczenia nieudana");
                MessageBox.Show("Kalibracja nieudana. Poczekaj aż silnik skończy działanie!", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button_zmiana_ustawien_pomiary_Click(this, EventArgs.Empty);
            }
        }



        private void button_zapisz_ustawienia_pomiary_Click(object sender, EventArgs e)
        {
            //KONTROLA BŁĘDÓW
            if (status_multimetr == false) //jeśli nie połączono się jeszcze z multimetrem
            {
                listbox_historia_add("Próba zapisania ustawień w zakładce 'Pomiary' nieudana z powodu braku połączenia z multimetrem.");
                MessageBox.Show("Nie można przejść dalej. Połącz się najpierw z multimetrem w zakładce 'Komunikacja'.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button_zmiana_ustawien_pomiary_Click(this, EventArgs.Empty);
                return; //wyjdź z funkcji button_zapisz_ustawienia_pomiary_Click
            }
            if (skalibrowany_czuj_przem == 999) //jeśli nieskalibrowano czujnika przemieszczenia
            {
                listbox_historia_add("Próba zapisania ustawień w zakładce 'Pomiary' nieudana z powodu nieskalibrowania czujnika przemieszczenia.");
                MessageBox.Show("Nie można zapisać ustawień. Czujnik przemieszczenia nie został poprawnie skalibrowany.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button_zmiana_ustawien_pomiary_Click(this, EventArgs.Empty);
                return; //wyjdź z funkcji button_zapisz_ustawienia_pomiary_Click
            }

            if (stan_comboBox_rodzaj_pomiaru == 2) //jeśli wybrano pomiar rezystancji
            {
                czy_odlaczono_zasilanie = MessageBox.Show("Wybrano metodę pomiaru rezystancji. Czy upewniłeś się, że na pewno odłączyłeś przewody doprowadzające zasilanie do mostka?", "Uwaga!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (czy_odlaczono_zasilanie == DialogResult.No)
                {
                    button_zmiana_ustawien_pomiary_Click(this, EventArgs.Empty);
                    listbox_historia_add("Użytkownik zrezygnował z zapisania ustawień z powodu nieodłączenia przewodów zasilających mostek.");
                    return; //wyjdź z funkcji button_zapisz_ustawienia_pomiary_Click
                }
                else
                {
                    //kontynuuj proces zapisywania ustawień i przejdź dalej
                }
            }
            

            comboBox_rodzaj_pomiaru.Enabled = false;
            button_zapisz_ustawienia_pomiary.Enabled = false;
            pictureBox_silnik_gora_pom.Enabled = false;
            pictureBox_silnik_dol_pom.Enabled = false;
            numericUpDown_dlugosc_ruchu.Enabled = false;
            numericUpDown_czulosc.Enabled = false;
            textBox_odczyt_czuj_przem_V.Enabled = false;
            button_kalibruj.Enabled = false;
            button_zmiana_ustawien_pomiary.Enabled = true;
            label_zapisano_ustawienia_pomiary.Visible = true;
            timer_pomiary_ustawienia.Enabled = false;
            panel_pomiary_pomiary.Enabled = true;
            timer_pomiary_pomiary.Enabled = true;
            czy_zapisano_ustawienia = true;
            polecenie_ugnij_belke = false;

            switch (stan_comboBox_rodzaj_pomiaru)
            {
                case 0: //2V
                    chart1.Titles.Clear();
                    title = chart1.Titles.Add("Charakterystyka napięcia w przekątnej mostka w funkcji ugięcia belki");
                    title.Font = new System.Drawing.Font("Calibri Light", 12, FontStyle.Regular);
                    title.Alignment = ContentAlignment.MiddleCenter;
                    chart1.ChartAreas[0].AxisY.Title = "U [mV]";
                    label_jednostka_multimetr.Text = "mV";
                    dataGridView1.Columns[2].HeaderText = "U [mV]";
                    break;

                case 1: //4V
                    chart1.Titles.Clear();
                    title = chart1.Titles.Add("Charakterystyka napięcia w przekątnej mostka w funkcji ugięcia belki");
                    title.Font = new System.Drawing.Font("Calibri Light", 12, FontStyle.Regular);
                    title.Alignment = ContentAlignment.MiddleCenter;
                    chart1.ChartAreas[0].AxisY.Title = "U [mV]";
                    label_jednostka_multimetr.Text = "mV";
                    dataGridView1.Columns[2].HeaderText = "U [mV]";
                    break;

                case 2: //OFF (rezystancja)
                    chart1.Titles.Clear();
                    title = chart1.Titles.Add("Charakterystyka rezystancji tensometru w funkcji ugięcia belki");
                    title.Font = new System.Drawing.Font("Calibri Light", 12, FontStyle.Regular);
                    title.Alignment = ContentAlignment.MiddleCenter;
                    chart1.ChartAreas[0].AxisY.Title = "R [Ω]";
                    label_jednostka_multimetr.Text = "Ω";
                    dataGridView1.Columns[2].HeaderText = "R [Ω]";
                    break;
            }
            listbox_historia_add("Pomyślnie zapisano ustawienia w zakładce 'Pomiary'.");

        }

        private void button_zmiana_ustawien_pomiary_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0) //jeśli tabela nie jest pusta
                zmiana_ust_czy_kasujemy_postep = MessageBox.Show("Spowoduje to usunięcie wszystkich danych z wykresu oraz z tabeli.\n\nCzy chcesz kontynuować?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);


            if (dataGridView1.Rows.Count == 0 || zmiana_ust_czy_kasujemy_postep == DialogResult.Yes) //jeśli tabela jest pusta lub użytkownik zgodził się na usunięcie dotychczasowych pomiarów
            {
                czyszczenie_wykresu_oraz_tabeli();
                chart1.Series["Multimetr"].Label = "";
                chart1.Series["Multimetr"].Points.AddXY(0, 0); //punkt na początek (0,0)
                skalibrowany_czuj_przem = 999; //ustawienie braku kalibracji czujnika przemieszczenia
                label_zapisano_ustawienia_pomiary.Visible = false;
                panel_pomiary_pomiary.Enabled = false;
                panel_pomiary_ustawienia.Enabled = true;
                comboBox_rodzaj_pomiaru.Enabled = true;
                button_zapisz_ustawienia_pomiary.Enabled = true;
                pictureBox_silnik_gora_pom.Enabled = true;
                pictureBox_silnik_dol_pom.Enabled = true;
                numericUpDown_dlugosc_ruchu.Enabled = true;
                numericUpDown_czulosc.Enabled = true;
                numericUpDown_dlugosc_ruchu.Value = 50; //początkowe ustawienie długości ruchu na 50%
                textBox_odczyt_czuj_przem_V.Enabled = true;
                button_kalibruj.Enabled = true;
                timer_pomiary_ustawienia.Enabled = true;
                timer_pomiary_pomiary.Enabled = false;
                czy_zapisano_ustawienia = false;
                reset_ukl_ster();
                textBox_odczyt_czuj_przem_V.BackColor = SystemColors.Control;
                textBox_odczyt_czuj_przem_V.TextAlign = HorizontalAlignment.Left;

                //Opcje wykresu i tabeli - wartości początkowe
                chart1.Titles.Clear();
                title = chart1.Titles.Add("Charakterystyka rezystancji tensometru w funkcji ugięcia belki lub\nCharakterystyka napięcia w przekątnej mostka w funkcji ugięcia belki");
                title.Font = new System.Drawing.Font("Calibri Light", 12, FontStyle.Regular);
                title.Alignment = ContentAlignment.MiddleCenter;
                chart1.ChartAreas[0].AxisX.Title = "Ugięcie belki [mm]";
                chart1.ChartAreas[0].AxisY.Title = "R [Ω] lub U [mV]";
                dataGridView1.Columns[2].HeaderText = "Multimetr";


                listbox_historia_add("Wybrano opcję zmiany ustawień w zakładce 'Pomiary'.");
            }
            else
            {
                //Nie zdecydowano się na zmianę ustawień z powodu chęci zostawienia postępu na wykresie/w tabeli
                //Nie rób nic
            }

        }




        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        // ********* CZĘŚĆ 2 - POMIARY **********
        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 

        private void timer_pomiary_pomiary_Tick(object sender, EventArgs e) //TIMER AKTYWUJĄCY SIĘ PO ZAPISANIU USTAWIEŃ I PRZEJŚCIU DO POMIARÓW
        {
            co_drugi_pomiar = !co_drugi_pomiar; //silnik porusza się co drugie wywołanie timera, bo inaczej odczyt odległości się zawiesza
            //ZAPYTANIA DO UKŁADU STEROWANIA
            try
            {
                switch (stan_comboBox_rodzaj_pomiaru)
                {
                    case 0: //2V
                        sp_ukl_ster.WriteLine("2");
                        break;

                    case 1: //4V
                        sp_ukl_ster.WriteLine("1");
                        break;

                    case 2: //OFF (rezystancja)
                        sp_ukl_ster.WriteLine("0");
                        break;
                }

                sp_ukl_ster.WriteLine("3");
                //textBox_odczyt_czuj_przem_V_pom.Text = sp_ukl_ster.ReadExisting(); //stara metoda - zapychał się bufor

                string odczyt = sp_ukl_ster.ReadExisting();
                string odczyt_v2 = "";
                double ugiecie_belki_mm = 0;

                try
                {
                    odczyt = (double.Parse(odczyt, System.Globalization.CultureInfo.InvariantCulture) * 5 / 1024).ToString(); //przeliczenie liczby od 0-1024 na napięcie
                

                if (!String.IsNullOrWhiteSpace(odczyt) && (odczyt.Length >= 4) && (odczyt[1] == ',')) //jeśli odczyt nie jest pusty/ ma minimum 4 znaki, a na drugiem miejscu jest przecinek
                {
                    textBox_odczyt_czuj_przem_V_pom.Text = odczyt.Substring(0, 4); //umieść w textBox pierwsze 4 znaki z bufora, np. 3.29
                    odczyt_v2 = odczyt; //zapisz w odczyt_v2 wszystkie znaki z bufora
                }

                    //PRZELICZANIE ODCZYTU Z CZUJNIKA PRZEMIESZCZENIA NA MILIMETRY
                    //czułość czuj. przem. ustawiana przez użytkownika aplikacji (numericUpDown_czulosc)
                    //double ugiecie_belki_mm = (double.Parse(odczyt, System.Globalization.CultureInfo.InvariantCulture) - skalibrowany_czuj_przem) / (Convert.ToDouble(numericUpDown_czulosc.Value) / 1000); //ugiecie belki [mm] = (aktualny odczyt [V] - skalibrowany odczyt [V])/czulosc

                        ugiecie_belki_mm =  (-1) * (double.Parse(odczyt_v2) - skalibrowany_czuj_przem) / (Convert.ToDouble(numericUpDown_czulosc.Value) / 1000); //ugiecie belki [mm] = (-1)*[(aktualny odczyt [V] - skalibrowany odczyt [V])/czulosc]

                textBox_odczyt_czuj_przem_mm.Text = ugiecie_belki_mm.ToString("0.00"); //wyświetlenie odczytu w textBox'ie
                }
                catch { } //nie wyrzucaj bledu jesli jeden z pomiarow sie nie powiedzie

                if (polecenie_ugnij_belke == true && co_drugi_pomiar == true && stop_motor == false) //jeśli kliknięto przycisk "ugnij belkę", jest to co drugi pomiar i nie kliknięto "STOP"
                {
                    //@@@ DOKŁADNOŚĆ DO USTAWIENIA!!!!!!!!
                    double dokladnosc = 0.15; //podana w mm
                    if (ugiecie_belki_mm < ustawione_ugiecie - dokladnosc) //jeśli belka jest niewystarczająco ugięta
                    {
                        //GetRandomNumber(25, 35)
                        for (int n = 0; n < 25; n++)
                        {
                            sp_ukl_ster.WriteLine("5"); //silnik dół
                        }
                    }
                    else if (ugiecie_belki_mm > ustawione_ugiecie + dokladnosc) //jeśli belka jest zbyt ugięta
                    {
                        for (int n = 0; n < 25; n++)
                        {
                            sp_ukl_ster.WriteLine("4"); //silnik góra
                        }
                    }
                    else  //jeśli belka jest dobrze ugięta
                    {
                        sp_ukl_ster.WriteLine("6"); //silnik OFF
                        polecenie_ugnij_belke = false; //zaprzestań uginanie belki w programie
                    }
                }
            } //koniec try

            catch
            {
                czy_zerwano_polaczenie_w_trakcie_wykonywania_pomiarow = true;
                timer_pomiary_pomiary.Enabled = false;
                listbox_historia_add("Połączenie z układem sterowania zostało zerwane w trakcie wykonywania pomiarów.");
                MessageBox.Show("Połączenie z układem sterowania zostało zerwane.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button_ukl_ster_rozlacz_Click(this, EventArgs.Empty);
            }



            //ZAPYTANIA DO MULTIMETRU

            //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            //myDmm.WriteString("*RST", true); //Reset the device //SPRAWDZIC CZY NIE POWODUJE BLEDOW
            //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

            try
            {
                //@@@ ODEBRANIE DANYCH Z MULTIMETRU - PÓŹNIEJ DO USTAWIENIA
                if (stan_comboBox_rodzaj_pomiaru == 0 || stan_comboBox_rodzaj_pomiaru == 1) //pomiar napięcia 2V lub 4V
                {
                    //textBox_odczyt_multimetr.Text = Math.Round(GetRandomNumber(0.1, 2.0), 3).ToString();

                    //Configure for DCV 0.2V range, 10uV resolution
                    //myDmm.WriteString("CONF:VOLT:DC 0.2, 0.00001", true);
                    //myDmm.WriteString("CONF:mVOLT:DC 0.01, 0.001", true);
                    myDmm.WriteString("CONF:VOLT:DC", true);
                    myDmm.WriteString("READ?", true);
                    string DCVResult = myDmm.ReadString();
                    string newDCVResult = DCVResult.Replace(',', '.'); //zamiana przecinka na kropke w stringu
                    double napiecie_odczyt = Convert.ToDouble(newDCVResult, System.Globalization.CultureInfo.InvariantCulture.NumberFormat);

                    double mvv = napiecie_odczyt * 1000;
                    textBox_odczyt_multimetr.Text = mvv.ToString("0.00"); //2 m. po przecinku
                    //textBox_odczyt_multimetr.Text = Math.Round(napiecie_odczyt,3).ToString(); //zaokrąglone do 3 miejsc po przecinku
                }
                if (stan_comboBox_rodzaj_pomiaru == 2) //pomiar rezystancji
                {
                    //textBox_odczyt_multimetr.Text = Math.Round(GetRandomNumber(249.90, 250.30), 2).ToString();
                    //Configure for OHM 2 wire 100 Ohm range, 100uOhm resolution
                    //myDmm.WriteString("CONF:RES 100, 0.0001", true);

                    //Configure for OHM 2 wire 250 Ohm range, 1mOhm resolution
                    myDmm.WriteString("CONF:RES 300, 0.0001", true);
                    myDmm.WriteString("READ?", true);
                    string Res2WResult = myDmm.ReadString();
                    string newRes2WResult = Res2WResult.Replace(',', '.'); //zamiana przecinka na kropke w stringu
                    double rezystancja_odczyt = Convert.ToDouble(newRes2WResult, System.Globalization.CultureInfo.InvariantCulture.NumberFormat);
                    //textBox_odczyt_multimetr.Text = rezystancja_odczyt.ToString(); @@@ DOKLADNE
                    textBox_odczyt_multimetr.Text = rezystancja_odczyt.ToString("0.00"); //2 m. po przecinku
                }
            }
            catch
            {
                czy_zerwano_polaczenie_w_trakcie_wykonywania_pomiarow = true;
                timer_pomiary_pomiary.Enabled = false;
                listbox_historia_add("Połączenie z multimetrem zostało zerwane w trakcie wykonywania pomiarów.");
                MessageBox.Show("Połączenie z multimetrem zostało zerwane.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button_keysight_rozlacz_Click(this, EventArgs.Empty);
            }

        }

        private void button_ugnij_belke_Click(object sender, EventArgs e)
        {
            stop_motor = false;
            ustawione_ugiecie = Decimal.ToDouble(numericUpDown_ustawienie_ugiecia.Value);
            polecenie_ugnij_belke = true;
            listbox_historia_add("Wysłano polecenie ugięcia belki do układu sterowania do wartości:  " + ustawione_ugiecie + ".");
        }

        private void button_stop_motor_Click(object sender, EventArgs e)
        {
            stop_motor = true;
            sp_ukl_ster.WriteLine("6");
            listbox_historia_add("Zatrzymano ręcznie silnik");
        }

        private void button_zapisz_pomiar_Click(object sender, EventArgs e)
        {
            //KONTROLA BŁĘDÓW
            if (double.Parse(textBox_odczyt_czuj_przem_mm.Text) <= maksymalne_zapisane_ugiecie_belki_mm()) //jeśli ugięcie belki jest mniejsze niż już wcześniej zapisane
            {
                listbox_historia_add("Próba zapisania pomiaru nieudana z powodu zbyt małego ugięcia belki w stosunku do wcześniejszych pomiarów.");
                MessageBox.Show("Aktualne ugięcie belki " + double.Parse(textBox_odczyt_czuj_przem_mm.Text) + " mm jest mniejsze lub równe wartości ugięcia zapisanej przy poprzednim pomiarze: " + maksymalne_zapisane_ugiecie_belki_mm() + " mm." + "\n\nZwiększ ugięcie belki i spróbuj ponownie.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; //wyjdź z funkcji button_zapisz_pomiar_Click
            }

            //pobranie wartości
            string ugiecie_belki_mm = textBox_odczyt_czuj_przem_mm.Text;
            string multimetr_odczyt = textBox_odczyt_multimetr.Text;


            string multimetr, jednostka; //zmienne do komunikatu
            if (stan_comboBox_rodzaj_pomiaru == 0 || stan_comboBox_rodzaj_pomiaru == 1) //zasilanie 2V lub 4V
            {
                multimetr = "U = ";
                jednostka = " mV";
            }
            else //pomiar rezystancji
            {
                multimetr = "R = ";
                jednostka = " Ω";
            }

            //czy chcesz zapisać pomiar? - komunikat
            czy_chcesz_zapisac_pomiar = MessageBox.Show("Ugięcie belki = " + ugiecie_belki_mm + " mm\n" + multimetr + multimetr_odczyt + jednostka + "\n\nCzy chcesz zapisać powyższe wartości?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (czy_chcesz_zapisac_pomiar == DialogResult.Yes) //użytkownik chce zapisać pomiar
            {
                if (lp == 1) //jeśli rysowanie wykresu dopiero się zaczyna (lp ==1)
                {
                    foreach (var series in chart1.Series)
                    {
                        series.Points.Clear(); //wyczyść wykres
                    }
                }

                //ta część będzie wykonywana jeśli już coś jest na wykresie

                //rysowanie obu osi nie rozpoczyna się od wartości '0'
                chart1.ChartAreas[0].AxisY.IsStartedFromZero = false;
                chart1.ChartAreas[0].AxisX.IsStartedFromZero = false;

                //dodaj do tabeli liczbe porzadkowa, ugiecie belki, odczyt z multimetru
                dataGridView1.Rows.Add(lp.ToString(), ugiecie_belki_mm, multimetr_odczyt);
                listbox_historia_add("Zapisano pomiar: " + lp.ToString() + ", " + ugiecie_belki_mm + ", " + multimetr_odczyt);

                //zapisanie wartości z tabeli do zmiennych typu doble wartX, wartY oraz ustawienie liczb po przecinku
                double wartX = float.Parse(dataGridView1.Rows[lp - 1].Cells[1].Value.ToString());
                double wartY = float.Parse(dataGridView1.Rows[lp - 1].Cells[2].Value.ToString());
                wartX = Math.Round(wartX, 2);
                wartY = Math.Round(wartY, 3);

                chart1.Series["Multimetr"].Points.AddXY(wartX, wartY); //dodaj punkt na wykresie
                //chart1.Series["Multimetr"].Points[lp - 1].Label = "(" + String.Format("{0:0.00}", wartX) + "; " + String.Format("{0:0.000}", wartY) + ")"; //stare współrzędne punktów (nie działało do końca)

                //jest już coś na wykresie, więc umożliw eksport wykresu/tabeli
                button_eksportuj_excel.Enabled = true;
                button_eksportuj_wykres.Enabled = true;

                lp++; //zwiększająca się liczba porządkowa
            }
            else
            {
                //użytkownik nie chce zapisać tego pomiaru
                //nie rób nic
            }
        }

        private void zaznacz_wspolrzedne_na_wykresie(string ans) //zaznaczenie współrzędnych punktów na wykresie
        {
            if (dataGridView1.Rows.Count == 0) //jeśli nie ma pomiarów
            {
                if (ans == "tak")
                    chart1.Series["Multimetr"].Points[0].Label = "(0,00; 0,00)"; //punkt 0,0
                else
                    chart1.Series["Multimetr"].Points[0].Label = ""; //pusto
            }

            //jeśli są już zapisane jakieś pomiary
            for (int n = 0; n < dataGridView1.Rows.Count; n++) //dla każdego pomiaru jaki już istnieje
            {
                if (ans == "tak")
                    chart1.Series["Multimetr"].Points[n].Label = "(" + String.Format("{0:0.00}", dataGridView1.Rows[n].Cells[1].Value) + "; " + String.Format("{0:0.000}", dataGridView1.Rows[n].Cells[2].Value) + ")"; //współrzędne ON
                else
                    chart1.Series["Multimetr"].Points[n].Label = ""; //pusto
            }

        }

        private void timer_wspolrzedne_Tick(object sender, EventArgs e) //timer, który aktywuje/dezaktywuje widoczność współrzędnych punktów co zadany czas
        {

            if (checkBox_wspolrzedne.Checked == true) //ON
                zaznacz_wspolrzedne_na_wykresie("tak");
            else if (checkBox_wspolrzedne.Checked == false) //OFF
                zaznacz_wspolrzedne_na_wykresie("nie");
        }

        private void button_pomiary_wyczysc_Click(object sender, EventArgs e) //USUŃ ZAPISANE DOTYCHCZAS POMIARY
        {
            if (dataGridView1.Rows.Count != 0) //jeśli tabela nie jest pusta
                czyszczenie_wykresu_i_tabeli_czy_kasujemy_postep = MessageBox.Show("Czynność ta spowoduje usunięcie wszystkich danych z wykresu oraz z tabeli.\n\n Czy chcesz kontynuować?", "Komunikat", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (dataGridView1.Rows.Count == 0 || czyszczenie_wykresu_i_tabeli_czy_kasujemy_postep == DialogResult.Yes) //jeśli nie ma żadnych pomiarów lub użytkownik zgodził się na usunięcie dotychczasowych pomiarów
            {
                czyszczenie_wykresu_oraz_tabeli();
                chart1.Series["Multimetr"].Label = "";
                chart1.Series["Multimetr"].Points.AddXY(0, 0); //punkt na początek (0,0)
                listbox_historia_add("Wyczyszczono tabelę i wykres.");
            }
        }





        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 
        // ********* CZĘŚĆ 3 - EKSPORT DANYCH DO EXCELA LUB WYKRES PLIKU GRAFICZNEGO **********
        // ********* ********* ********* ********* ********* ********* ********* ********* ********* 

        private void button_eksportuj_excel_Click(object sender, EventArgs e)
        {
            
            //deklaracje obiektów
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {
                worksheet = workbook.ActiveSheet;

                //ustalanie nazw arkusza
                switch (stan_comboBox_rodzaj_pomiaru)
                {
                    case 0: //2V
                        worksheet.Name = "Pomiary_cw_tensometry_Uzas=2V";
                        break;

                    case 1: //4V
                        worksheet.Name = "Pomiary_cw_tensometry_Uzas=4V";
                        break;

                    case 2: //OFF (rezystancja)
                        worksheet.Name = "Pomiary_cw_tensometry_Uzas=0V";
                        break;

                    default:
                        worksheet.Name = "Pomiary";
                        break;

                }

                //zapisywanie wszystkich danych do komórek w exceli
                int cellRowIndex = 1;
                int cellColumnIndex = 1;

                for (int i = -1; i <= dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++) //0
                    {
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = float.Parse(dataGridView1.Rows[i].Cells[j].Value.ToString());
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveDialog.FilterIndex = 1;
                saveDialog.Title = "Zapisywanie tabeli jako...";

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Eksport tabeli zakończony powodzeniem.");
                    listbox_historia_add("Eksport tabeli zakończony powodzeniem.");
                }
            }

            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message); //error box
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }

        } //koniec button_eksportuj_excel_Click




        private void button_eksportuj_wykres_Click(object sender, EventArgs e)
        {
            try
            {
                //Check if chart has at least one enabled series with points
                if (chart1.Series.Any(s => s.Enabled && s.Points.Count > 0))
                {
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "Image Files|*.png;";
                    save.Filter = "Png Image (.png)|*.bmp|Gif Image (.gif)|*.gif|JPEG Image (.jpeg)|*.jpeg|Bitmap Image (.bmp)|*.png|Tiff Image (.tiff)|*.tiff";
                    save.Title = "Zapisywanie wykresu jako...";
                    save.DefaultExt = ".png";
                    if (save.ShowDialog() == DialogResult.OK)
                    {
                        //chart1.Size = new Size(1240,880); //powiększenie wykresu chwilę przed zapisaniem go (nie przyniosło pożądanych efektów)
                        var imgFormats = new Dictionary<string, ChartImageFormat>()
            {
                {".bmp", ChartImageFormat.Bmp},
                {".gif", ChartImageFormat.Gif},
                {".jpg", ChartImageFormat.Jpeg},
                {".jpeg", ChartImageFormat.Jpeg},
                {".png", ChartImageFormat.Png},
                {".tiff", ChartImageFormat.Tiff},
            };
                        var fileExt = System.IO.Path.GetExtension(save.FileName).ToString().ToLower();

                        if (imgFormats.ContainsKey(fileExt))
                        {
                            chart1.SaveImage(save.FileName, imgFormats[fileExt]);
                            MessageBox.Show("Eksport wykresu do pliku graficznego zakończony powodzeniem.");
                            listbox_historia_add("Eksport wykresu do pliku graficznego zakończony powodzeniem.");
                        }
                        else
                        {
                            throw new Exception(String.Format("Tylko formaty plików graficznych '{0}' są wspierane", string.Join(", ", imgFormats.Keys)));
                        }
                    }
                }
                else
                {
                    throw new Exception("Brak danych do eksportu!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("SaveChartAsImage()", ex.Message);
            }
            //chart1.Size = new Size(620,440); //powrót do normalnego rozmiaru wykresu (nie przyniosło pożądanych efektów)
        }


        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------










        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      OBSŁUGA ZAKŁADKI "HISTORIA"
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        private void button_clear_listBox_Click(object sender, EventArgs e) //CZYSZCZENIE LISTY "HISTORIA"
        {
            listBox_historia.Items.Clear();
        }

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------










        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      OBSŁUGA ZAKŁADKI "INFORMACJE"
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        //funkcja do wyłączenia textbox'a
        private void textBox_GotFocus(object sender, EventArgs e)
        {
            ((TextBox)sender).Parent.Focus();
        }

        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------










        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //
        //      OBSŁUGA ZAKŁADKI "FUNKCJE POMOCNICZE"
        //
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        //GENERATOR LICZB LOSOWYCH ZMIENNOPRZECINKOWYCH Z DANEGO ZAKRESU (POTRZEBNE DLA TESTÓW APLIKACJI)
        public double GetRandomNumber(double minimum, double maximum)
        {
            Random random = new Random();
            return random.NextDouble() * (maximum - minimum) + minimum;
        }


        private void timer_data_czas_Tick(object sender, EventArgs e) //TIMER DO WYŚWIETLANIA AKTUALNEJ DATY I GODZINY
        {
            label_data_i_czas.Text = System.DateTime.Now.ToString();
        }


        private void reset_ukl_ster() //WYŁĄCZENIE NAPIĘĆ 2V I 4V ORAZ WYŁĄCZENIE SILNIKA
        {
            if (status_ukl_ster == true)
            {
                try
                {
                    sp_ukl_ster.WriteLine("0"); //napięcia mostka OFF
                    sp_ukl_ster.WriteLine("6"); //silniki OFF
                }
                catch { } //pusto
            }
        }

        private void autoscroll_listbox(string ktory_listbox) //AUTOSCROLLOWANIE LISTBOX'A
        {
            switch (ktory_listbox)
            {
                case "listBox_odp_multimetr":
                    listBox_odp_multimetr.SelectedIndex = listBox_odp_multimetr.Items.Count - 1;
                    listBox_odp_multimetr.SelectedIndex = -1;
                    break;

                case "listBox_historia":
                    listBox_historia.SelectedIndex = listBox_historia.Items.Count - 1;
                    listBox_historia.SelectedIndex = -1;
                    break;
            }
        }


        private void listbox_historia_add(string wiadomosc) //UMIESZCZENIE INFORMACJI W ZAKŁADCE 'HISTORIA' ORAZ AUTOMATYCZNY AUTOSCROLL
        {
            listBox_historia.Items.Add(System.DateTime.Now.ToString() + " // " + wiadomosc);
            autoscroll_listbox("listBox_historia");
        }

       
        private void listbox_odp_multimetr_add(string wiadomosc) //UMIESZCZENIE ODP. MULTIMETRU W ZAKŁADCE 'TESTOWANIE' ORAZ AUTOMATYCZNY AUTOSCROLL
        {
            listBox_odp_multimetr.Items.Add(wiadomosc);
            autoscroll_listbox("listBox_odp_multimetr");
        }


        private void czyszczenie_wykresu_oraz_tabeli() //CZYSZCZENIE WYKRESU I TABELI W ZAKŁADCE 'POMIARY'
        {
            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            lp = 1;

            button_eksportuj_excel.Enabled = false;
            button_eksportuj_wykres.Enabled = false;
        }


        private double maksymalne_zapisane_ugiecie_belki_mm() //WYZNACZANIE NAJWIĘKSZEJ DOTYCHCZAS ZAPISANEJ WARTOŚCI W TABELI DATAGRIDVIEW1
        {
            double max_ugiecie_mm = -999; //deklaracja zmiennej

            for (int n = 0; n < dataGridView1.Rows.Count; n++) //dla wszystkich wierszy ugięcia belki
            {
                if (double.Parse(dataGridView1.Rows[n].Cells[1].Value.ToString()) > max_ugiecie_mm) //szukanie największej odczytanie wartości
                {
                    max_ugiecie_mm = double.Parse(dataGridView1.Rows[n].Cells[1].Value.ToString()); //zapisanie jej do zmiennej max_ugiece_mm
                }
            }

            return max_ugiecie_mm; //zwrócenie największej wartości
        }








        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


    } //klasa Form1
} //namespace



