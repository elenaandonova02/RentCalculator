using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Rentu_Kalkulator_Aktuarstvo_2021
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        int x = 0;
        int n = 0;
        int k = 0;
        int i = 0;
        int p = 0;
        double E = 0.00;
        double Nx = 0.00;
        double Dx = 0.00;
        double Nx2 = 0.00;
        double Sx = 0.00;
        double Sx2 = 0.00;
        double a = 0.00;
        double M = 0.00;
        double R = 0.00;
        
      //  double q = 0.00;

        public Form1()
        {
            InitializeComponent();
            anticipativna1.Enabled = false;
            dekurzivna1.Enabled = false;
            privremena1.Enabled = false;
            dozivotna1.Enabled = false;
            odlozena1.Enabled = false;
            neposredna1.Enabled = false;
            podatoci.Enabled = false;

            anticipativna.Enabled = false;
            dekurzivna.Enabled = false;
            privremena.Enabled = false;
            dozivotna.Enabled = false;
            odlozena.Enabled = false;
            neposredna.Enabled = false;
            groupBox1.Enabled = false;

            anticipativna2.Enabled = false;
            dekurzivna2.Enabled = false;
            privremena2.Enabled = false;
            dozivotna2.Enabled = false;
            aritmeticka.Enabled = false;
            geometriska.Enabled = false;
            presmetajProm.Enabled = false;
            groupBox2.Enabled = false;
            zgolemuva.Enabled = false;
            namaluva.Enabled = false;
            procentProm.Enabled = false;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //otvaranje na excel
            OpenFile();
        }

        //funkcija OpenFile za otvaranje excel
        public void OpenFile()
        {
            //kreiranje nov objekt za otvaranje na excel
            Excel excel = new Excel(@"C:\Users\ELENA\source\repos\Rentu_Kalkulator_Aktuarstvo_2021\Mortality_table_in_MKD_Tablici_na_smrtnost_MK.xlsx", 1);
            //MessageBox.Show(excel.ReadCell(0, 0));
        }

        //calculate_Click funkcija za kopce koe so klik presmetuva GODISHNA RENTA
        private void calculate_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel(@"C:\Users\ELENA\source\repos\Rentu_Kalkulator_Aktuarstvo_2021\Mortality_table_in_MKD_Tablici_na_smrtnost_MK.xlsx", 1);

            if (maski_pol.Checked)
            {
                if (anticipativna1.Checked)
                {
                    if (neposredna1.Checked)
                    {
                        if (dozivotna1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                i = x + 4;

                                Nx = double.Parse(excel.getCellValue(x + 4, 7));
                                Dx = double.Parse(excel.getCellValue(x + 4, 6));
                                //ovie dvete dole message boxovi se za proverka prvo ti go dava decimalniot broj po Nx, a posle po Dx
                                // posle toa izleguva presmetkata
                                //MessageBox.Show(excel.getCellValue(x + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x + 4, 6));

                                a = (Nx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                //else
                   // MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna1.Checked)
                {
                    if (neposredna1.Checked)
                    {
                        if (dozivotna1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                x = x + 1;
                                Nx = double.Parse(excel.getCellValue(x + 4, 7));
                                Dx = double.Parse(excel.getCellValue(x + 3, 6));
                                //MessageBox.Show(excel.getCellValue(x + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x + 3, 6));

                                a = (Nx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
               // else
                  //  MessageBox.Show("Немате внесено доволно податоци!");

                if (anticipativna1.Checked)
                {
                    if (odlozena1.Checked)
                    {
                        if (dozivotna1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || odlozuvanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                k = int.Parse(odlozuvanje.Text);
                                //x = x + k;
                                Nx = double.Parse(excel.getCellValue(x + k + 4, 7));
                                Dx = double.Parse(excel.getCellValue(x + 4, 6));
                                //MessageBox.Show(excel.getCellValue(x + k + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x + 4, 6));

                                a = (Nx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
              //  else
              //      MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna1.Checked)
                {
                    if (odlozena1.Checked)
                    {
                        if (dozivotna1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || odlozuvanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                k = int.Parse(odlozuvanje.Text);
                                x = x + k + 1;
                                Nx = double.Parse(excel.getCellValue(x + 4, 7));
                                Dx = double.Parse(excel.getCellValue(x - 17, 6));
                                //MessageBox.Show(excel.getCellValue(x + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x - 17, 6));

                                a = (Nx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
               // else
                    //MessageBox.Show("Немате внесено доволно податоци!");

                if (anticipativna1.Checked)
                {
                    if (neposredna1.Checked)
                    {
                        if (privremena1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || period_na_primanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                n = int.Parse(period_na_primanje.Text);
                                // x = x + n + 1;
                                Nx = double.Parse(excel.getCellValue(x + 4, 7));
                                Dx = double.Parse(excel.getCellValue(x + 4, 6));
                                //MessageBox.Show(excel.getCellValue(x + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x + 4, 6));
                                Nx2 = double.Parse(excel.getCellValue(x + n + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x + n + 4, 7));

                                a = (Nx * 100000 - (Nx2 * 100000)) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                //else
                   // MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna1.Checked)
                {
                    if (neposredna1.Checked)
                    {
                        if (privremena1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || period_na_primanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                n = int.Parse(period_na_primanje.Text);
                                Nx = double.Parse(excel.getCellValue(x + 5, 7));
                                Dx = double.Parse(excel.getCellValue(x + 4, 6));
                                //MessageBox.Show(excel.getCellValue(x + 5, 7));
                                //MessageBox.Show(excel.getCellValue(x + 4, 6));
                                Nx2 = double.Parse(excel.getCellValue(x + n + 5, 7));
                                //MessageBox.Show(excel.getCellValue(x + n + 5, 7));

                                a = (Nx * 100000 - (Nx2 * 100000)) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
               // else
                   // MessageBox.Show("Немате внесено доволно податоци!");

                if (anticipativna1.Checked)
                {
                    if (odlozena1.Checked)
                    {
                        if (privremena1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || period_na_primanje.Text != "" || odlozuvanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                n = int.Parse(period_na_primanje.Text);
                                k = int.Parse(odlozuvanje.Text);
                                Nx = double.Parse(excel.getCellValue(x + k + 4, 7));
                                Dx = double.Parse(excel.getCellValue(x + 4, 6));
                                //MessageBox.Show(excel.getCellValue(x + k + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x + 4, 6));
                                Nx2 = double.Parse(excel.getCellValue(x + k + n + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x + k + n + 4, 7));

                                a = (Nx * 100000 - (Nx2 * 100000)) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
               // else
                  //  MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna1.Checked)
                {
                    if (odlozena1.Checked)
                    {
                        if (privremena1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || period_na_primanje.Text != "" || odlozuvanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                n = int.Parse(period_na_primanje.Text);
                                k = int.Parse(odlozuvanje.Text);
                                Nx = double.Parse(excel.getCellValue(x + k + 5, 7));
                                Dx = double.Parse(excel.getCellValue(x + 4, 6));
                                //MessageBox.Show(excel.getCellValue(x + k + 5, 7));
                                //MessageBox.Show(excel.getCellValue(x + 4, 6));
                                Nx2 = double.Parse(excel.getCellValue(x + k + n + 5, 7));
                                //MessageBox.Show(excel.getCellValue(x + k + n + 5, 7));

                                a = (Nx * 100000 - (Nx2 * 100000)) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
               // else
                  //  MessageBox.Show("Немате внесено доволно податоци!");
            }

            if (zenski_pol.Checked)
            {
                if (anticipativna1.Checked)
                {
                    if (neposredna1.Checked)
                    {
                        if (dozivotna1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                i = x + 4;

                                Nx = double.Parse(excel.getCellValue(x + 106, 7));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                //ovie dvete dole message boxovi se za proverka prvo ti go dava decimalniot broj po Nx, a posle po Dx
                                // posle toa izleguva presmetkata
                                //MessageBox.Show(excel.getCellValue(x + 106, 7));
                                //MessageBox.Show(excel.getCellValue(x + 106, 6));

                                a = (Nx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                //else
                  //  MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna1.Checked)
                {
                    if (neposredna1.Checked)
                    {
                        if (dozivotna1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                x = x + 1;
                                Nx = double.Parse(excel.getCellValue(x + 106, 7));
                                Dx = double.Parse(excel.getCellValue(x + 105, 6));
                                //MessageBox.Show(excel.getCellValue(x + 106, 7));
                                //MessageBox.Show(excel.getCellValue(x + 105, 6));

                                a = (Nx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
              //  else
                //    MessageBox.Show("Немате внесено доволно податоци!");

                if (anticipativna1.Checked)
                {
                    if (odlozena1.Checked)
                    {
                        if (dozivotna1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || odlozuvanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                k = int.Parse(odlozuvanje.Text);
                                //x = x + k;
                                Nx = double.Parse(excel.getCellValue(x + k + 4, 7));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                //MessageBox.Show(excel.getCellValue(x + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x - 16, 6));

                                a = (Nx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
               // else
                  //  MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna1.Checked)
                {
                    if (odlozena1.Checked)
                    {
                        if (dozivotna1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || odlozuvanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                k = int.Parse(odlozuvanje.Text);
                                //x = x + k + 1;
                                Nx = double.Parse(excel.getCellValue(x + k + 1 + 106, 7));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                //MessageBox.Show(excel.getCellValue(x + 4, 7));
                                //MessageBox.Show(excel.getCellValue(x - 17, 6));

                                a = (Nx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
               // else
                   // MessageBox.Show("Немате внесено доволно податоци!");


                if (anticipativna1.Checked)
                {
                    if (neposredna1.Checked)
                    {
                        if (privremena1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || period_na_primanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                n = int.Parse(period_na_primanje.Text);
                                // x = x + n + 1;
                                Nx = double.Parse(excel.getCellValue(x + 106, 7));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                //MessageBox.Show(excel.getCellValue(x + 106, 7));
                                //MessageBox.Show(excel.getCellValue(x + 106, 6));
                                Nx2 = double.Parse(excel.getCellValue(x + n + 106, 7));
                                //MessageBox.Show(excel.getCellValue(x + n + 106, 7));

                                a = ((Nx * 100000) - (Nx2 * 100000)) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
               // else
                 //   MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna1.Checked)
                {
                    if (neposredna1.Checked)
                    {
                        if (privremena1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || period_na_primanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                n = int.Parse(period_na_primanje.Text);
                                Nx = double.Parse(excel.getCellValue(x + 106, 7));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                //MessageBox.Show(excel.getCellValue(x + 107, 7));
                                //MessageBox.Show(excel.getCellValue(x + 106, 6));
                                Nx2 = double.Parse(excel.getCellValue(x + n + 107, 7));
                                //MessageBox.Show(excel.getCellValue(x + n + 107, 7));

                                a = ((Nx * 100000) - (Nx2 * 100000)) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                //else
                   // MessageBox.Show("Немате внесено доволно податоци!");

                if (anticipativna1.Checked)
                {
                    if (odlozena1.Checked)
                    {
                        if (privremena1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || period_na_primanje.Text != "" || odlozuvanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                n = int.Parse(period_na_primanje.Text);
                                k = int.Parse(odlozuvanje.Text);
                                Nx = double.Parse(excel.getCellValue(x + k + 106, 7));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                //MessageBox.Show(excel.getCellValue(x + k + 106, 7));
                                //MessageBox.Show(excel.getCellValue(x + 106, 6));
                                Nx2 = double.Parse(excel.getCellValue(x + k + n + 106, 7));
                                //MessageBox.Show(excel.getCellValue(x + k + n + 106, 7));

                                a = (Nx * 100000 - (Nx2 * 100000)) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                //else
                  //  MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna1.Checked)
                {
                    if (odlozena1.Checked)
                    {
                        if (privremena1.Checked)
                        {
                            if (renta.Text != "" || vozrast.Text != "" || period_na_primanje.Text != "" || odlozuvanje.Text != "")
                            {
                                x = int.Parse(vozrast.Text);
                                R = double.Parse(renta.Text);
                                n = int.Parse(period_na_primanje.Text);
                                k = int.Parse(odlozuvanje.Text);
                                Nx = double.Parse(excel.getCellValue(x + k + 107, 7));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                //MessageBox.Show(excel.getCellValue(x + k + 107, 7));
                                //MessageBox.Show(excel.getCellValue(x + 106, 6));
                                Nx2 = double.Parse(excel.getCellValue(x + k + n + 107, 7));
                                //MessageBox.Show(excel.getCellValue(x + k + n + 107, 7));

                                a = (Nx * 100000 - (Nx2 * 100000)) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
               // else
                  //  MessageBox.Show("Немате внесено доволно податоци!");
            }
        }
    
        private void privremena1_CheckedChanged(object sender, EventArgs e)
        {
            period_na_primanje.Enabled = true;
        }

        private void dozivotna1_CheckedChanged(object sender, EventArgs e)
        {
            period_na_primanje.Enabled = false;
        }

        private void odlozena1_CheckedChanged(object sender, EventArgs e)
        {
            odlozuvanje.Enabled = true;
        }

        private void neposredna1_CheckedChanged(object sender, EventArgs e)
        {
            odlozuvanje.Enabled = false;
        }

        private void zenski_pol_CheckedChanged(object sender, EventArgs e)
        {
            anticipativna1.Enabled = true;
            dekurzivna1.Enabled = true;
            privremena1.Enabled = true;
            dozivotna1.Enabled = true;
            odlozena1.Enabled = true;
            neposredna1.Enabled = true;
            podatoci.Enabled = true;
        }

        private void maski_pol_CheckedChanged(object sender, EventArgs e)
        {
            anticipativna1.Enabled = true;
            dekurzivna1.Enabled = true;
            privremena1.Enabled = true;
            dozivotna1.Enabled = true;
            odlozena1.Enabled = true;
            neposredna1.Enabled = true;
            podatoci.Enabled = true;
        }



        private void btnMiza_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel(@"C:\Users\ELENA\source\repos\Rentu_Kalkulator_Aktuarstvo_2021\Mortality_table_in_MKD_Tablici_na_smrtnost_MK.xlsx", 1);
            //odreduvame koj tip na renta e odbran za da ja presmetame soodvetnata
            //toa go pravime so check button na radio button (dali e pritisnato)
            //odvojuvame za maz i za zena zaradi razlicni vrednosti i pozicii vo tabeli
            int x = int.Parse(vozrastPar.Text);
            int renta = int.Parse(rentaPar.Text);


            //double Miza = 0.00;
            //double Dx = 0.00;
            //double Nx = 0.00;


            if (masko.Checked == true)
            {
                //presmetki za maz
                //Varijanti na Renta, od 1-8
                if (anticipativna.Checked == true && neposredna.Checked == true && dozivotna.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "") 
                    { 
                        //Varijanta 1
                        int m = Convert.ToInt32(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 4, 7));
                        Dx = double.Parse(excel.getCellValue(x + 4, 6));

                        double ax = (Nx * 100000) / (Dx * 100000);
                        double axm = ax - ((m - 1) / (2 * m));
                        M = R * axm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                        // MessageBox.Show("      Nx=" + Nx + "       Dx=" + Dx);
                    }
                }

                else if (dekurzivna.Checked == true && neposredna.Checked == true && dozivotna.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "")
                    {

                        //Varijanta 2
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 4, 7));
                        Dx = double.Parse(excel.getCellValue(x + 4, 6));
                        double ax = (Nx * 100000) / (Dx * 100000);

                        int m = Convert.ToInt32(brPrimanjaPar.Text);
                        double axm = ax - ((m - 1) / (2 * m));
                        M = R * axm;

                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }

                else if (anticipativna.Checked == true && odlozena.Checked == true && dozivotna.Checked == true)
                {
                    //ToDo -> CONNECT EXCEL SHEET TO KOMUTATIVNI BROEVI
                    //Varijanta 3
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || odlozuvanjePar.Text != "")
                    {
                        int k = int.Parse(odlozuvanjePar.Text);
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        double Nxk = double.Parse(excel.getCellValue(x + 4 + k, 7));

                        Dx = double.Parse(excel.getCellValue(x + 4, 6));

                        double Dxk = double.Parse(excel.getCellValue(x + 4 + k, 6));

                        double kax = (Nxk * 100000) / (Dx * 100000);
                        double kEx = (Dxk * 100000) / (Dx * 100000);

                        double kaxm = kax - ((m - 1) / (2 * m)) * kEx;
                        M = R * kaxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }
                else if (dekurzivna.Checked == true && odlozena.Checked == true && dozivotna.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || odlozuvanjePar.Text != "")
                    {
                        //Varijanta 4
                        int k = int.Parse(odlozuvanjePar.Text);
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);

                        Dx = double.Parse(excel.getCellValue(x + 4, 6));

                        double Nxk1 = double.Parse(excel.getCellValue(x + 4 + k + 1, 7));
                        double Dxk = double.Parse(excel.getCellValue(x + 4 + k, 6));

                        double kax = (Nxk1 * 100000) / (Dx * 100000);
                        double kEx = (Dxk * 100000) / (Dx * 100000);

                        double kaxm = kax - ((m - 1) / (2 * m)) * kEx;

                        M = R * kaxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }


                else if (anticipativna.Checked == true && neposredna.Checked == true && privremena.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || periodNaPrimanjePar.Text != "")
                    {
                        //Varijanta 5
                        int m = int.Parse(brPrimanjaPar.Text);
                        int n = int.Parse(periodNaPrimanjePar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);

                        Nx = double.Parse(excel.getCellValue(x + 4, 7));
                        Dx = double.Parse(excel.getCellValue(x + 4, 6));

                        double Nxn = double.Parse(excel.getCellValue(x + 4 + n, 7));

                        double nax = ((Nx * 100000) - (Nxn * 100000)) / (Dx * 100000);


                        double naxm = nax - (1 - (Dx + n / Dx)) * ((m - 1) / (2 * m));
                        M = R * naxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }
                else if (dekurzivna.Checked == true && neposredna.Checked == true && privremena.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || periodNaPrimanjePar.Text != "")
                    {
                        //Varijanta 6
                        int m = int.Parse(brPrimanjaPar.Text);
                        int n = int.Parse(periodNaPrimanjePar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 4, 7));
                        Dx = double.Parse(excel.getCellValue(x + 4, 6));
                        double Nxn1 = double.Parse(excel.getCellValue(x + 4 + n + 1, 7));
                        double Dxn = double.Parse(excel.getCellValue(x + 4 + n, 6));

                        double nax = ((Nx * 100000) - (Nxn1 * 100000)) / (Dx * 100000);


                        double naxm = nax - (1 - (Dxn / Dx)) * ((m - 1) / (2 * m));
                        M = R * naxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }


                else if (anticipativna.Checked == true && odlozena.Checked == true && privremena.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || periodNaPrimanjePar.Text != "" || odlozuvanjePar.Text != "")
                    {
                        //Varijanta 7
                        int k = int.Parse(odlozuvanjePar.Text);
                        int m = int.Parse(brPrimanjaPar.Text);
                        int n = int.Parse(periodNaPrimanjePar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 4, 7));
                        Dx = double.Parse(excel.getCellValue(x + 4, 6));

                        double Nxk = double.Parse(excel.getCellValue(x + 4 + k, 7));
                        double Nxkn = double.Parse(excel.getCellValue(x + 4 + k + n, 7));
                        double Dxk = double.Parse(excel.getCellValue(x + 4 + k, 6));
                        double Dxkn = double.Parse(excel.getCellValue(x + 4 + k + n, 6));


                        double knax = ((Nxk * 100000) - (Nxkn * 100000)) / (Dx * 100000);


                        double knaxm = knax - ((Dxk * 100000) / (Dx * 100000)) - ((Dxkn * 100000) / (Dx * 100000)) * ((m - 1) / (2 * m));
                        M = R * knaxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }
                else if (dekurzivna.Checked == true && odlozena.Checked == true && privremena.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || periodNaPrimanjePar.Text != "" || odlozuvanjePar.Text != "")
                    {
                        //ToDo -> CONNECT EXCEL SHEET TO KOMUTATIVNI BROEVI
                        //Varijanta 8
                        int k = int.Parse(odlozuvanjePar.Text);
                        int m = int.Parse(brPrimanjaPar.Text);
                        int n = int.Parse(periodNaPrimanjePar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 4, 7));
                        Dx = double.Parse(excel.getCellValue(x + 4, 6));

                        double Nxk1 = double.Parse(excel.getCellValue(x + 4 + k + 1, 7));
                        double Nxkn1 = double.Parse(excel.getCellValue(x + 4 + k + n + 1, 7));
                        double Dxk = double.Parse(excel.getCellValue(x + 4 + k, 6));
                        double Dxkn = double.Parse(excel.getCellValue(x + 4 + k + n, 6));


                        double knax = ((Nxk1 * 100000) - (Nxkn1 * 100000)) / (Dx * 100000);


                        double knaxm = knax - ((Dxk * 100000) / (Dx * 100000)) - ((Dxkn * 100000) / (Dx * 100000)) * ((m - 1) / (2 * m));
                        M = R * knaxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }

            }

            else if (zensko.Checked == true)
            {
                //presmetki za zena
                //Varijanti na Renta, od 1-8
                if (anticipativna.Checked == true && neposredna.Checked == true && dozivotna.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "")
                    {
                        //Varijanta 1
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 106, 7));
                        Dx = double.Parse(excel.getCellValue(x + 106, 6));

                        double ax = (Nx * 100000) / (Dx * 100000);
                        double axm = ax - ((m - 1) / (2 * m));
                        M = R * axm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                        // MessageBox.Show("      Nx=" + Nx + "       Dx=" + Dx);
                    }

                }

                else if (dekurzivna.Checked == true && neposredna.Checked == true && dozivotna.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "")
                    {

                        //Varijanta 2
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 106, 7));
                        Dx = double.Parse(excel.getCellValue(x + 106, 6));
                        double ax = (Nx * 100000) / (Dx * 100000);
                        double axm = ax - ((m - 1) / (2 * m));
                        M = R * axm;

                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }

                else if (anticipativna.Checked == true && odlozena.Checked == true && dozivotna.Checked == true)
                {
                    //ToDo -> CONNECT EXCEL SHEET TO KOMUTATIVNI BROEVI
                    //Varijanta 3
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || odlozuvanjePar.Text != "")
                    {
                        int k = int.Parse(odlozuvanjePar.Text);
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        double Nxk = double.Parse(excel.getCellValue(x + 106 + k, 7));

                        Dx = double.Parse(excel.getCellValue(x + 106, 6));

                        double Dxk = double.Parse(excel.getCellValue(x + 106 + k, 6));

                        double kax = (Nxk * 100000) / (Dx * 100000);
                        double kEx = (Dxk * 100000) / (Dx * 100000);

                        double kaxm = kax - ((m - 1) / (2 * m)) * kEx;
                        M = R * kaxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }
                else if (dekurzivna.Checked == true && odlozena.Checked == true && dozivotna.Checked == true)
                {
                    //Varijanta 4
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || odlozuvanjePar.Text != "")
                    {
                        int k = int.Parse(odlozuvanjePar.Text);
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        Dx = double.Parse(excel.getCellValue(x + 106, 6));

                        double Nxk1 = double.Parse(excel.getCellValue(x + 106 + k + 1, 7));
                        double Dxk = double.Parse(excel.getCellValue(x + 106 + k, 6));

                        double kax = (Nxk1 * 100000) / (Dx * 100000);
                        double kEx = (Dxk * 100000) / (Dx * 100000);

                        double kaxm = kax - ((m - 1) / (2 * m)) * kEx;

                        M = R * kaxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }


                else if (anticipativna.Checked == true && neposredna.Checked == true && privremena.Checked == true)
                {
                    //Varijanta 5
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || periodNaPrimanjePar.Text != "")
                    {
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        int n = int.Parse(periodNaPrimanjePar.Text);

                        Nx = double.Parse(excel.getCellValue(x + 106, 7));
                        Dx = double.Parse(excel.getCellValue(x + 106, 6));

                        double Nxn = double.Parse(excel.getCellValue(x + 106 + n, 7));

                        double nax = ((Nx * 100000) - (Nxn * 100000)) / (Dx * 100000);

                        double naxm = nax - (1 - (Dx + n / Dx)) * ((m - 1) / (2 * m));
                        M = R * naxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }
                else if (dekurzivna.Checked == true && neposredna.Checked == true && privremena.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || periodNaPrimanjePar.Text != "")
                    {
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        int n = int.Parse(periodNaPrimanjePar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 106, 7));
                        Dx = double.Parse(excel.getCellValue(x + 106, 6));
                        double Nxn1 = double.Parse(excel.getCellValue(x + 106 + n + 1, 7));
                        double Dxn = double.Parse(excel.getCellValue(x + 106 + n, 6));

                        double nax = ((Nx * 100000) - (Nxn1 * 100000)) / (Dx * 100000);


                        double naxm = nax - (1 - (Dxn / Dx)) * ((m - 1) / (2 * m));
                        M = R * naxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }


                else if (anticipativna.Checked == true && odlozena.Checked == true && privremena.Checked == true)
                {
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || periodNaPrimanjePar.Text != "" || odlozuvanjePar.Text != "")
                    {
                        //Varijanta 7
                        int k = int.Parse(odlozuvanjePar.Text);
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        int n = int.Parse(periodNaPrimanjePar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 106, 7));
                        Dx = double.Parse(excel.getCellValue(x + 106, 6));

                        double Nxk = double.Parse(excel.getCellValue(x + 106 + k, 7));
                        double Nxkn = double.Parse(excel.getCellValue(x + 106 + k + n, 7));
                        double Dxk = double.Parse(excel.getCellValue(x + 106 + k, 6));
                        double Dxkn = double.Parse(excel.getCellValue(x + 106 + k + n, 6));


                        double knax = ((Nxk * 100000) - (Nxkn * 100000)) / (Dx * 100000);


                        double knaxm = knax - ((Dxk * 100000) / (Dx * 100000)) - ((Dxkn * 100000) / (Dx * 100000)) * ((m - 1) / (2 * m));
                        M = R * knaxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                    }
                }
            
                else if (dekurzivna.Checked == true && odlozena.Checked == true && privremena.Checked == true)
                {
                    //ToDo -> CONNECT EXCEL SHEET TO KOMUTATIVNI BROEVI
                    //Varijanta 8
                    if (vozrastPar.Text != "" || brPrimanjaPar.Text != "" || rentaPar.Text != "" || periodNaPrimanjePar.Text != "" || odlozuvanjePar.Text != "")
                    {
                        int k = int.Parse(odlozuvanjePar.Text);
                        int m = int.Parse(brPrimanjaPar.Text);
                        x = int.Parse(vozrastPar.Text);
                        R = double.Parse(rentaPar.Text);
                        int n = int.Parse(periodNaPrimanjePar.Text);
                        Nx = double.Parse(excel.getCellValue(x + 106, 7));
                        Dx = double.Parse(excel.getCellValue(x + 106, 6));

                        double Nxk1 = double.Parse(excel.getCellValue(x + 106 + k + 1, 7));
                        double Nxkn1 = double.Parse(excel.getCellValue(x + 106 + k + n + 1, 7));
                        double Dxk = double.Parse(excel.getCellValue(x + 106 + k, 6));
                        double Dxkn = double.Parse(excel.getCellValue(x + 106 + k + n, 6));


                        double knax = ((Nxk1 * 100000) - (Nxkn1 * 100000)) / (Dx * 100000);


                        double knaxm = knax - ((Dxk * 100000) / (Dx * 100000)) - ((Dxkn * 100000) / (Dx * 100000)) * ((m - 1) / (2 * m));
                        M = R * knaxm;
                        MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");

                    }   
                }
            }
        }
        private void privremena_CheckedChanged(object sender, EventArgs e)
        {
            periodNaPrimanjePar.Enabled = true;
            odlozuvanjePar.Enabled = true;
        }

        private void dozivotna_CheckedChanged(object sender, EventArgs e)
        {
            periodNaPrimanjePar.Enabled = false;
        }

        private void odlozena_CheckedChanged(object sender, EventArgs e)
        {
            odlozuvanjePar.Enabled = true;
        }

        private void neposredna_CheckedChanged(object sender, EventArgs e)
        {
            odlozuvanjePar.Enabled = false;
        }
        private void masko_CheckedChanged(object sender, EventArgs e)
        {
            //enable radio buttons
            anticipativna.Enabled = true;
            dekurzivna.Enabled = true;
            privremena.Enabled = true;
            dozivotna.Enabled = true;
            odlozena.Enabled = true;
            neposredna.Enabled = true;
            btnMiza.Enabled = true;
            groupBox1.Enabled = true;

        }

        private void zensko_CheckedChanged(object sender, EventArgs e)
        {
            //enable radio buttons
            anticipativna.Enabled = true;
            dekurzivna.Enabled = true;
            privremena.Enabled = true;
            dozivotna.Enabled = true;
            odlozena.Enabled = true;
            neposredna.Enabled = true;
            btnMiza.Enabled = true;
            groupBox1.Enabled = true;
        }

        private void presmetajProm_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel(@"C:\Users\ELENA\source\repos\Rentu_Kalkulator_Aktuarstvo_2021\Mortality_table_in_MKD_Tablici_na_smrtnost_MK.xlsx", 1);

            if (maz.Checked)
            {
                if (anticipativna2.Checked)
                {
                    if (aritmeticka.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (rentaProm.Text != "" || vozrastProm.Text != "")
                            {
                                x = int.Parse(vozrastProm.Text);
                                R = double.Parse(rentaProm.Text);
                                i = x + 4;

                                Sx = double.Parse(excel.getCellValue(x + 4, 9));
                                Dx = double.Parse(excel.getCellValue(x + 4, 6));
                                //ovie dvete dole message boxovi se za proverka prvo ti go dava decimalniot broj po Nx, a posle po Dx
                                // posle toa izleguva presmetkata
                                //MessageBox.Show(excel.getCellValue(x + 4, 9));
                                //MessageBox.Show(excel.getCellValue(x + 4, 6));

                                a = (Sx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                //else
                // MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna2.Checked)
                {
                    if (aritmeticka.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (rentaProm.Text != "" || vozrastProm.Text != "")
                            {
                                x = int.Parse(vozrastProm.Text);
                                R = double.Parse(rentaProm.Text);
                                x = x + 1;
                                Sx = double.Parse(excel.getCellValue(x + 4, 9));
                                Dx = double.Parse(excel.getCellValue(x + 3, 6));
                                //MessageBox.Show(excel.getCellValue(x + 4, 9));
                                //MessageBox.Show(excel.getCellValue(x + 3, 6));

                                a = (Sx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                // else
                //  MessageBox.Show("Немате внесено доволно податоци!");

                if (anticipativna2.Checked)
                {
                    if (aritmeticka.Checked)
                    {
                        if (privremena2.Checked)
                        {
                            if (rentaProm.Text != "" || vozrastProm.Text != "" || periodProm.Text != "")
                            {
                                x = int.Parse(vozrastProm.Text);
                                R = double.Parse(rentaProm.Text);
                                n = int.Parse(periodProm.Text);
                                //x = x + k;
                                Sx = double.Parse(excel.getCellValue(x + 4, 9));
                                Sx2 = double.Parse(excel.getCellValue(x + n + 4, 9));
                                Dx = double.Parse(excel.getCellValue(x + 4, 6));
                                Nx = double.Parse(excel.getCellValue(x + n + 4, 7));

                                //MessageBox.Show(excel.getCellValue(x + n + 4, 9));
                                //MessageBox.Show(excel.getCellValue(x + 4, 6));
                                //MessageBox.Show(excel.getCellValue(x + n + 4, 7));

                                a = ((Sx * 100000) - (Sx2 * 100000) - (n * (Nx * 100000))) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                //  else
                //      MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna2.Checked)
                {
                    if (aritmeticka.Checked)
                    {
                        if (privremena2.Checked)
                        {
                            if (rentaProm.Text != "" || vozrastProm.Text != "" || periodProm.Text != "")
                            {
                                x = int.Parse(vozrastProm.Text);
                                R = double.Parse(rentaProm.Text);
                                n = int.Parse(periodProm.Text);
                                //x = x + k;
                                Sx = double.Parse(excel.getCellValue(x + 4 + 1, 9));
                                Sx2 = double.Parse(excel.getCellValue(x + n + 4 + 1, 9));
                                Dx = double.Parse(excel.getCellValue(x + 4, 6));
                                Nx = double.Parse(excel.getCellValue(x + n + 4 + 1, 7));

                                //MessageBox.Show(excel.getCellValue(x + n + 4 + 1, 9));
                                //MessageBox.Show(excel.getCellValue(x + 4, 6));
                                //MessageBox.Show(excel.getCellValue(x + n + 4 + 1, 7));

                                a = ((Sx * 100000) - (Sx2 * 100000) - (n * (Nx * 100000))) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                // else
                //MessageBox.Show("Немате внесено доволно податоци!");
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (zgolemuva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 4, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 4 + 1, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 4, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 4, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + 1, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 4, 6));

                                    a = ((Nx * 100000) + (E * (Sx * 100000))) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (namaluva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    E = p / 100;
                                    
                                    Nx = double.Parse(excel.getCellValue(x + 4, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 4 + 1, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 4, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 4, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + 1, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 4, 6));

                                    a = ((Nx * 100000) - (E * (Sx * 100000))) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (zgolemuva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 4 + 1, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 4 + 2, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 4, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 4 + 1, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + 2, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 4, 6));

                                    a = ((Nx * 100000) + (E * (Sx * 100000))) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (namaluva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 4 + 1, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 4 + 2, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 4, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 4 + 1, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + 2, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 4, 6));

                                    a = ((Nx * 100000) - E * (Sx * 100000)) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (privremena2.Checked)
                        {
                            if (zgolemuva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "" || periodProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    n = int.Parse(periodProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 4, 7));
                                    Nx2 = double.Parse(excel.getCellValue(x + 4 + n, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 4 + 1, 9));
                                    Sx2 = double.Parse(excel.getCellValue(x + 4 + n, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 4, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 4, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + n, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + 1, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + n, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 4, 6));


                                    a = (((Nx * 100000) - (Nx2 * 100000)) + (E * ((Sx * 100000) - (Sx2 * 100000) - ((n-1) * Nx2)))) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (namaluva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "" || periodProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    n = int.Parse(periodProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 4 + 1, 7));
                                    Nx2 = double.Parse(excel.getCellValue(x + 4 + n + 1, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 4 + 2, 9));
                                    Sx2 = double.Parse(excel.getCellValue(x + 4 + n + 1, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 4, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 4 + 1, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + n + 1, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + 2, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 4 + n + 1, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 4, 6));


                                    a = (((Nx * 100000) - (Nx2 * 100000)) - (E * ((Sx * 100000) - (Sx2 * 100000) - ((n - 1) * Nx2)))) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
            }
            if (zena.Checked)
            {
                if (anticipativna2.Checked)
                {
                    if (aritmeticka.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (rentaProm.Text != "" || vozrastProm.Text != "")
                            {
                                x = int.Parse(vozrastProm.Text);
                                R = double.Parse(rentaProm.Text);
                                i = x + 4;

                                Sx = double.Parse(excel.getCellValue(x + 106, 9));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                //ovie dvete dole message boxovi se za proverka prvo ti go dava decimalniot broj po Nx, a posle po Dx
                                // posle toa izleguva presmetkata
                                //MessageBox.Show(excel.getCellValue(x + 106, 9));
                                //MessageBox.Show(excel.getCellValue(x + 106, 6));

                                a = (Sx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                //else
                // MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna2.Checked)
                {
                    if (aritmeticka.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (rentaProm.Text != "" || vozrastProm.Text != "")
                            {
                                x = int.Parse(vozrastProm.Text);
                                R = double.Parse(rentaProm.Text);
                                x = x + 1;
                                Sx = double.Parse(excel.getCellValue(x + 106, 9));
                                Dx = double.Parse(excel.getCellValue(x + 105, 6));
                                //MessageBox.Show(excel.getCellValue(x + 106, 9));
                                //MessageBox.Show(excel.getCellValue(x + 105, 6));

                                a = (Sx * 100000) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                // else
                //  MessageBox.Show("Немате внесено доволно податоци!");

                if (anticipativna2.Checked)
                {
                    if (aritmeticka.Checked)
                    {
                        if (privremena2.Checked)
                        {
                            if (rentaProm.Text != "" || vozrastProm.Text != "" || periodProm.Text != "")
                            {
                                x = int.Parse(vozrastProm.Text);
                                R = double.Parse(rentaProm.Text);
                                n = int.Parse(periodProm.Text);
                                //x = x + k;
                                Sx = double.Parse(excel.getCellValue(x + 106, 9));
                                Sx2 = double.Parse(excel.getCellValue(x + n + 106, 9));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                Nx = double.Parse(excel.getCellValue(x + n + 106, 7));

                                //MessageBox.Show(excel.getCellValue(x + n + 106, 9));
                                //MessageBox.Show(excel.getCellValue(x + 106, 6));
                                //MessageBox.Show(excel.getCellValue(x + n + 106, 7));

                                a = ((Sx * 100000) - (Sx2 * 100000) - (n * (Nx * 100000))) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                //  else
                //      MessageBox.Show("Немате внесено доволно податоци!");

                if (dekurzivna2.Checked)
                {
                    if (aritmeticka.Checked)
                    {
                        if (privremena2.Checked)
                        {
                            if (rentaProm.Text != "" || vozrastProm.Text != "" || periodProm.Text != "")
                            {
                                x = int.Parse(vozrastProm.Text);
                                R = double.Parse(rentaProm.Text);
                                n = int.Parse(periodProm.Text);
                                //x = x + k;
                                Sx = double.Parse(excel.getCellValue(x + 106 + 1, 9));
                                Sx2 = double.Parse(excel.getCellValue(x + n + 106 + 1, 9));
                                Dx = double.Parse(excel.getCellValue(x + 106, 6));
                                Nx = double.Parse(excel.getCellValue(x + n + 106 + 1, 7));

                                //MessageBox.Show(excel.getCellValue(x + n + 106 + 1, 9));
                                //MessageBox.Show(excel.getCellValue(x + 106, 6));
                                //MessageBox.Show(excel.getCellValue(x + n + 106 + 1, 7));

                                a = ((Sx * 100000) - (Sx2 * 100000) - (n * (Nx * 100000))) / (Dx * 100000);
                                M = a * R;
                                MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                            }
                        }
                    }
                }
                // else
                //MessageBox.Show("Немате внесено доволно податоци!");
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (zgolemuva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 106, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 106 + 1, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 106, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 106, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + 1, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 106, 6));

                                    a = ((Nx * 100000) + E * (Sx * 100000)) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (namaluva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 106, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 106 + 1, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 106, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 106, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + 1, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 106, 6));

                                    a = ((Nx * 100000) - E * (Sx * 100000)) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (zgolemuva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 106 + 1, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 106 + 2, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 106, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 106 + 1, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + 2, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 106, 6));

                                    a = ((Nx * 100000) + E * (Sx * 100000)) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (namaluva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 106 + 1, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 106 + 2, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 106, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 106 + 1, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + 2, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 106, 6));

                                    a = ((Nx * 100000) - E * (Sx * 100000)) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (privremena2.Checked)
                        {
                            if (zgolemuva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "" || periodProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    n = int.Parse(periodProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 106, 7));
                                    Nx2 = double.Parse(excel.getCellValue(x + 106 + n, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 106 + 1, 9));
                                    Sx2 = double.Parse(excel.getCellValue(x + 106 + n, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 106, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 106, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + n, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + 1, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + n, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 106, 6));


                                    a = ((Nx * 100000) - (Nx2 * 100000) + E * ((Sx * 100000) - (Sx2 * 100000) - ((n - 1) * Nx2))) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
                if (anticipativna2.Checked)
                {
                    if (geometriska.Checked)
                    {
                        if (dozivotna2.Checked)
                        {
                            if (namaluva.Checked)
                            {
                                if (rentaProm.Text != "" || vozrastProm.Text != "" || procentProm.Text != "" || periodProm.Text != "")
                                {
                                    x = int.Parse(vozrastProm.Text);
                                    R = double.Parse(rentaProm.Text);
                                    p = int.Parse(procentProm.Text);
                                    n = int.Parse(periodProm.Text);
                                    E = p / 100;

                                    Nx = double.Parse(excel.getCellValue(x + 106 + 1, 7));
                                    Nx2 = double.Parse(excel.getCellValue(x + 106 + n + 1, 7));
                                    Sx = double.Parse(excel.getCellValue(x + 106 + 2, 9));
                                    Sx2 = double.Parse(excel.getCellValue(x + 106 + n + 1, 9));
                                    Dx = double.Parse(excel.getCellValue(x + 106, 6));

                                    //MessageBox.Show(excel.getCellValue(x + 106 + 1, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + n + 1, 7));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + 2, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 106 + n + 1, 9));
                                    //MessageBox.Show(excel.getCellValue(x + 106, 6));


                                    a = ((Nx * 100000) - (Nx2 * 100000) - E * ((Sx * 100000) - (Sx2 * 100000) - ((n - 1) * Nx2))) / (Dx * 100000);
                                    M = a * R;
                                    MessageBox.Show("Треба да се уплати " + M + " денари за примање на рента од " + R + " денари");
                                }
                            }
                        }
                    }
                }
            }
        }

        private void dozivotna2_CheckedChanged(object sender, EventArgs e)
        {
            periodProm.Enabled = false;
        }

        private void privremena2_CheckedChanged(object sender, EventArgs e)
        {
            periodProm.Enabled = true;
        }

        private void geometriska_CheckedChanged(object sender, EventArgs e)
        {
            procentProm.Enabled = true;
        }

        private void aritmeticka_CheckedChanged(object sender, EventArgs e)
        {
            procentProm.Enabled = false;
        }
        private void maz_CheckedChanged(object sender, EventArgs e)
        {
            //enable radio buttons
            anticipativna2.Enabled = true;
            dekurzivna2.Enabled = true;
            privremena2.Enabled = true;
            dozivotna2.Enabled = true;
            presmetajProm.Enabled = true;
            aritmeticka.Enabled = true;
            geometriska.Enabled = true;
            groupBox2.Enabled = true;
            procentProm.Enabled = true;
            

        }

        private void zena_CheckedChanged(object sender, EventArgs e)
        {
            //enable radio buttons
            anticipativna2.Enabled = true;
            dekurzivna2.Enabled = true;
            privremena2.Enabled = true;
            dozivotna2.Enabled = true;
            aritmeticka.Enabled = true;
            geometriska.Enabled = true;
            presmetajProm.Enabled = true;
            groupBox2.Enabled = true;
            procentProm.Enabled = true;
        }

        private void geometriska_CheckedChanged_1(object sender, EventArgs e)
        {
            zgolemuva.Enabled = true;
            namaluva.Enabled = true;
        }

        private void aritmeticka_CheckedChanged_1(object sender, EventArgs e)
        {
            zgolemuva.Enabled = false;
            namaluva.Enabled = false;
        }

        private void promenlivaRenta_Click(object sender, EventArgs e)
        {

        }
    }
}
