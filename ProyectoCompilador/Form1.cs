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
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Office.Interop.Excel;

namespace ProyectoCompilador
{
    public partial class Form1 : Form
    {
        string server = "Data Source = DESKTOP-VCEO21F; Initial Catalog= SistemasProgramacion; Integrated Security = True ";
        SqlConnection conectar = new SqlConnection();
        private string[] vectorPalabrasReservadas;
        private int[,] matriz;
        private int renglon;
        private string wlinea;
        private int direccion;
        private bool espalreservada;
        private int estado;
        private string token, wsalida;
        private int posicion;
        private char caracter;
        private int columna;
        int a = 0;
        int selectedId;
        string ruta, rutaV, usuario;
        public System.Data.DataTable CargarCombo()
        {
            conectar.ConnectionString = server;
            conectar.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter("CargarComboBox",conectar);
            dataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
            System.Data.DataTable dataTable = new System.Data.DataTable();
            dataAdapter.Fill(dataTable);
            return dataTable;
            
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            MessageBox.Show("Bienvenido "+usuario);
            // Código dentro del evento Form1_Load
            conectar.ConnectionString = server;
            cmbxLenguajes.DataSource=CargarCombo();
            cmbxLenguajes.DisplayMember = "nombre";
            cmbxLenguajes.ValueMember = "idlenguaje";            
            //SqlDataReader registro = cmd.ExecuteReader();
            //while (registro.Read())
            //{
            //    cmbxLenguajes.Items.Add(registro["nombre"].ToString());
            //    a = Convert.ToInt32(registro["idlenguaje"]);
            //    //MessageBox.Show(a.ToString());
            //}

            //    int b = cmbxLenguajes.SelectedIndex = a-1;
            //MessageBox.Show(b.ToString());
            //conectar.Close();
        }

        private void cmbxLenguajes_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            if (cmbxLenguajes.SelectedValue != null && int.TryParse(cmbxLenguajes.SelectedValue.ToString(), out selectedId))
            {
                if (selectedId == 1)
                {
                    ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                    matriz = new int[10,14];
                    vectorPalabrasReservadas = new string[37];
                    try
                    {
                        string renglon;
                        string[] datosrenglon;
                        int r = 0;
                        using (StreamReader lector = new StreamReader(ruta))
                        {
                            while (!lector.EndOfStream)
                            {
                                renglon = lector.ReadLine();
                                datosrenglon = renglon.Split(',');
                                for (int c = 0; c < datosrenglon.Length; c++)
                                {
                                    matriz[r, c] = Convert.ToInt32(datosrenglon[c]);

                                    
                                }

                                r++;
                            }
                        }
                        rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_VisualBasic.csv";
                        string renglonV;
                        using (StreamReader lector = new StreamReader(rutaV))
                        { 
                            renglonV = lector.ReadLine();
                        }
                        vectorPalabrasReservadas = renglonV.Split(',');
                        for(int i = 0; i<vectorPalabrasReservadas.Length; i++)
                        {
                            ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error"+ex.Message);
                    }
                }
                              // Utilizar el valor de selectedId
            }
            else if (selectedId == 2)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[10, 16];
                vectorPalabrasReservadas = new string[30];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_Basic.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }

            }
            else if (selectedId == 3)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[12, 20];
                vectorPalabrasReservadas = new string[31];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_C.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }
            }
            else if (selectedId == 4)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[12, 20];
                vectorPalabrasReservadas = new string[74];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_Clipper.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }
            }
            else if (selectedId == 5)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[12, 20];
                vectorPalabrasReservadas = new string[27];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_Cobol.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }
            }
            else if (selectedId == 6)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[16, 25];
                vectorPalabrasReservadas = new string[37];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_Dbase.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }
            }
            else if (selectedId == 7)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[50, 50];
                vectorPalabrasReservadas = new string[30];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_Fortran.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }
            }
            else if (selectedId == 8)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[12, 23];
                vectorPalabrasReservadas = new string[50];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_Java.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }
            }
            else if (selectedId == 9)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[13, 23];
                vectorPalabrasReservadas = new string[45];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_Pascal.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }
            }
            else if (selectedId == 10)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[11, 14];
                vectorPalabrasReservadas = new string[29];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_Python.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }
            }
            else if (selectedId == 11)
            {
                ruta = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Matrices de estados\" + selectedId.ToString() + ".csv";
                matriz = new int[9, 19];
                vectorPalabrasReservadas = new string[20];
                try
                {
                    string renglon;
                    string[] datosrenglon;
                    int r = 0;
                    using (StreamReader lector = new StreamReader(ruta))
                    {
                        while (!lector.EndOfStream)
                        {
                            renglon = lector.ReadLine();
                            datosrenglon = renglon.Split(',');
                            for (int c = 0; c < datosrenglon.Length; c++)
                            {
                                matriz[r, c] = Convert.ToInt32(datosrenglon[c]);


                            }

                            r++;
                        }
                    }
                    rutaV = @"C:\Users\alber\source\repos\ProyectoCompilador\ProyectoCompilador\bin\Debug\Palabras Reservadas\Palabras_R_Visual_FoxPro.csv";
                    string renglonV;
                    using (StreamReader lector = new StreamReader(rutaV))
                    {
                        renglonV = lector.ReadLine();
                    }
                    vectorPalabrasReservadas = renglonV.Split(',');
                    for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
                    {
                        ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error" + ex.Message);
                }
            }
        }

        private void cmbxLenguajes_MouseClick(object sender, MouseEventArgs e)
        {
        }
        public Form1(string usuario)
        {
            InitializeComponent();
            this.usuario = usuario;
            //leerMatrizEstados("matriz.csv");
            //leerPalabrasReservadas("Palabras_R_VisualBasic.csv");
            //for (int i = 0; i < vectorPalabrasReservadas.Length; i++)
            //{
            //    ListPreservada.Items.Add(vectorPalabrasReservadas[i] + "");
            //}
        }

        


        private void leerPalabrasReservadas(string archivo)
        {
            string renglon;
            using (StreamReader lector = new StreamReader(archivo))
            {
                renglon = lector.ReadLine();
            }
            vectorPalabrasReservadas = renglon.Split(',');
        }


        private void leerMatrizEstados(string archivo)
        {
            string renglon;
            string[] datosrenglon;
            int r = 0;
            using (StreamReader lector = new StreamReader(archivo))
            {
                while (!lector.EndOfStream)
                {
                    renglon = lector.ReadLine();
                    datosrenglon = renglon.Split(',');
                    for (int c = 0; c < datosrenglon.Length; c++)
                    {
                        matriz[r, c] = Convert.ToInt32(datosrenglon[c]);
                    }
                    r++;
                }
            }
        }
        private void BuscarTokens()
        {
            string apoyo;
            estado = 0;
            token = "";
            posicion = 1;
            while (posicion <= wlinea.Length)
            {
                apoyo = wlinea.Substring(posicion - 1, 1); // Extrae un carácter de wlinea
                caracter = apoyo.FirstOrDefault();
                CalcularColumna();
                estado = matriz[estado, columna];
                if (estado >= 100)
                {
                    if (token.Length > 0)
                    {
                        ReconoceTokens();
                    }
                    else if (token.Length == 0) // Únicamente caracteres especiales de un carácter
                    {
                        token = token + caracter;
                        ReconoceTokens();
                    }
                    estado = 0;
                    token = "";
                }
                else
                {
                    if (estado != 0) // Mientras sea diferente de 0, sigue agregando caracteres
                    {
                        token = token + caracter;
                    }
                }
                posicion++;
            }
            if (token.Length > 0)
            {
                caracter = ' ';
                CalcularColumna();
                estado = matriz[estado, columna];
                ReconoceTokens();
            }
        }

        private void CalcularColumna()
        {
            if (caracter >= 'A' && caracter <= 'Z' || caracter >= 'a' && caracter <= 'z')
            {
                columna = 0;
            }
            else if (caracter == ' ' || caracter == ' ')
            {
                columna = 2;
            }
            else if (caracter >= '0' && caracter <= '9')
            {
                columna = 1;
            }
            else if (caracter == '.')
            {
                columna = 3;
            }
            else if (caracter == '"')
            {
                columna = 4;
            }
            else if (caracter == '\'')
            {
                columna = 5;
            }
            else if (caracter == '+')
            {
                columna = 6;
            }
            else if (caracter == '-')
            {
                columna = 7;
            }
            else if (caracter == '*')
            {
                columna = 8;
            }
            else if (caracter == '/')
            {
                columna = 9;
            }
            else if (caracter == '<')
            {
                columna = 10;
            }
            else if (caracter == '>')
            {
                columna = 11;
            }
            else if (caracter == '=')
            {
                columna = 12;
            }
            else if (caracter == '_')
            {
                columna = 13;
            }
        }


        private void ReconoceTokens()
        {
            if (estado == 100)
            {
                espalreservada = false;
                BuscapalReservada();
                if (espalreservada)
                {
                    wsalida = token + "   PalReserv   ";//+ direccion.ToString();
                }
                else
                {
                    //Buscaidentificadores();
                    wsalida = token + " Ident  ";// + direccion.ToString();
                }
                posicion = posicion - 1; // Regresa una posición requirió de un delimitador
            }

            if (estado == 101)
            {
                //BuscarEnteras();
                wsalida = token + " Cte. Enteras ";// + direccion.ToString();
                posicion = posicion - 1;
            }
            else if (estado == 102)
            {
                //BuscarReales();
                wsalida = token + " Cte. Real";// + direccion.ToString();
                posicion = posicion - 1;
            }

            if (estado == 105)
            {
                wsalida = token + " Car. Esp";
            }
            else if (estado == 106)
            {
                wsalida = token + " Car. Esp";
            }
            else if (estado == 107)
            {
                wsalida = token + " Car. Esp";
            }
            else if (estado == 108)
            {
                wsalida = token + " Car. Esp";
            }
            else if (estado == 109)
            {
                wsalida = token + " Car. Esp";
                posicion = posicion - 1;
            }
            else if (estado == 110)
            {
                token = token + caracter;
                wsalida = token + " Car. Esp";
            }
            else if (estado == 111)
            {
                wsalida = token + " Car. Esp";
                posicion = posicion - 1;
            }
            else if (estado == 112)
            {
                token = token + caracter;
                wsalida = token + " Car. Esp";
            }
            else if (estado == 113)
            {
                wsalida = token + " Car. Esp";
            }
            else if (estado == 104)
            {
                token = token + caracter;
                wsalida = token + " Comment";
            }

            if (estado == 103)
            {
                token = token + caracter;
                //BuscarStrings();
                wsalida = token + " Cte. String "; //+ direccion.ToString();
            }

            ListSalida.Items.Add(wsalida);
        }


        //private void Buscaidentificadores()
        //{
        //    bool encontro = false;
        //    int renglon2 = 0;
        //    while (!encontro && renglon2 < ListIds.Items.Count)
        //    {
        //        ListIds.SelectedIndex = renglon2;
        //        if (token.ToUpper() == ListIds.Text.ToUpper())
        //        {
        //            encontro = true;
        //            direccion = renglon2;
        //        }
        //        renglon2 = renglon2 + 1;
        //    }
        //    if (!encontro)
        //    {
        //        ListIds.Items.Add(token);
        //        direccion = ListIds.Items.Count - 1;
        //    }
        //}


        private void BuscapalReservada()
        {
            int renglon2 = 0;
            direccion = -1;
            while (!espalreservada && renglon2 < vectorPalabrasReservadas.Length)
            {
                if (token.ToUpper() == vectorPalabrasReservadas[renglon2].ToUpper())
                {
                    espalreservada = true;
                    direccion = renglon2;
                }
                renglon2 = renglon2 + 1;
            }
        }


        private void PictureBox3_Click(object sender, EventArgs e)
        {
            ListEntrada.Items.Clear();
            ListSalida.Items.Clear();
            
        }


        private void PictureBox4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //private void BuscarEnteras()
        //{
        //    bool encontro = false;
        //    int renglon2 = 0;
        //    while (!encontro && renglon2 < ListEnteras.Items.Count)
        //    {
        //        ListEnteras.SelectedIndex = renglon2;
        //        if (token.ToUpper() == ListEnteras.Text.ToUpper())
        //        {
        //            encontro = true;
        //            direccion = renglon2;
        //        }
        //        renglon2 = renglon2 + 1;
        //    }
        //    if (!encontro)
        //    {
        //        direccion = ListEnteras.Items.Add(token);
        //        direccion = ListEnteras.Items.Count - 1;
        //    }
        //}

        //private void BuscarReales()
        //{
        //    bool encontro = false;
        //    int renglon2 = 0;
        //    while (!encontro && renglon2 < ListReales.Items.Count)
        //    {
        //        ListReales.SelectedIndex = renglon2;
        //        if (token.ToUpper() == ListReales.Text.ToUpper())
        //        {
        //            encontro = true;
        //            direccion = renglon2;
        //        }
        //        renglon2 = renglon2 + 1;
        //    }
        //    if (!encontro)
        //    {
        //        direccion = ListReales.Items.Add(token);
        //        direccion = ListReales.Items.Count - 1;
        //    }
        //}


        private void Button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialogoGuardar = new SaveFileDialog();

            if (dialogoGuardar.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            string archivo = dialogoGuardar.FileName;
            using (StreamWriter sw = new StreamWriter(archivo))
            {
                foreach (var item in ListSalida.Items)
                {
                    if (item.ToString() != "")
                    {
                        sw.WriteLine(item.ToString());
                    }
                }
            }
        }

        private void btnTXT_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string query = "Select nombre from Lenguajes where idlenguaje=@Id_lenguaje";
            SqlCommand cmd = new SqlCommand(query, conectar);
            cmd.Parameters.AddWithValue("@Id_lenguaje", selectedId.ToString());
            conectar.Close();
            conectar.Open();
            object resultado = cmd.ExecuteScalar();

            DateTime fechaHoraActual = DateTime.Now;

            string nombreArchivo = $"Output_{resultado}_{usuario}_{fechaHoraActual:yyy_MM_dd_HH_mm_ss}.csv";
            string rutaCompletaArchivo = Path.Combine(@"C:\\Users\\alber\\source\\repos\\ProyectoCompilador\\ProyectoCompilador\\bin\\Debug\\Outputs\\", nombreArchivo);
            using (StreamWriter writer = new StreamWriter(rutaCompletaArchivo))
            {
                foreach (var item in ListSalida.Items)
                {
                    writer.WriteLine(item.ToString());
                }
            }
            MessageBox.Show("Archivo creado existosamente");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string query = "Select nombre from Lenguajes where idlenguaje=@Id_lenguaje";
            SqlCommand cmd = new SqlCommand(query, conectar);
            cmd.Parameters.AddWithValue("@Id_lenguaje", selectedId.ToString());
            object resultado = cmd.ExecuteScalar();
            
            Excel.Application APPEXCEL = new Excel.Application();
            APPEXCEL.Visible = false;
            Excel.Workbook WOORKBOOK = APPEXCEL.Workbooks.Add();
            Excel.Worksheet WORKSHEET = WOORKBOOK.ActiveSheet as Excel.Worksheet;
            DateTime fechaHoraActual = DateTime.Now;

            string nombreArchivo = $"Output_{resultado}_{usuario}_{fechaHoraActual:yyy_MM_dd_HH_mm_ss}.xlsx";
            string rutaCompletaArchivo = Path.Combine(@"C:\\Users\\alber\\source\\repos\\ProyectoCompilador\\ProyectoCompilador\\bin\\Debug\\Outputs\\", nombreArchivo);
            
            for(int i = 0; i < ListSalida.Items.Count; i++)
            {
                string item = ListSalida.Items[i].ToString();
                WORKSHEET.Cells[i + 1,1] = item;
            }
            WOORKBOOK.SaveAs(rutaCompletaArchivo);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            conectar.Close();
            conectar.Open();
            string query = "Select nombre from Lenguajes where idlenguaje=@Id_lenguaje";
            SqlCommand cmd = new SqlCommand(query, conectar);
            cmd.Parameters.AddWithValue("@Id_lenguaje", selectedId.ToString());
            object resultado = cmd.ExecuteScalar();
           
            DateTime fechaHoraActual = DateTime.Now;

            string nombreArchivo = $"Output_{resultado}_{usuario}_{fechaHoraActual:yyy_MM_dd_HH_mm_ss}.txt";
            string rutaCompletaArchivo = Path.Combine(@"C:\\Users\\alber\\source\\repos\\ProyectoCompilador\\ProyectoCompilador\\bin\\Debug\\Outputs\\",nombreArchivo);
            using (StreamWriter writer = new StreamWriter(rutaCompletaArchivo))
            {
                foreach (var item in ListSalida.Items)
                {
                    writer.WriteLine(item.ToString());
                }
            }
            MessageBox.Show("Archivo creado existosamente");
            conectar.Close();
            string idusuario = ObtenerUsuario(usuario);
            
            int idusuarioreal = Convert.ToInt32(idusuario);
            conectar.Open();
            SqlCommand CMD = new SqlCommand("AgregarRegistro", conectar);
            CMD.CommandType = CommandType.StoredProcedure;
            CMD.Parameters.AddWithValue("@idusuario", idusuarioreal);
            CMD.Parameters.AddWithValue("@idlenguaje", selectedId);
            CMD.Parameters.AddWithValue("@fecha_hora", fechaHoraActual);
            CMD.Parameters.AddWithValue("@nombre_archivo", nombreArchivo);

            try
            {
                CMD.ExecuteNonQuery();
            }
            catch(SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                throw ex;
            }
            conectar.Close();
        }

        private void btnclean_Click(object sender, EventArgs e)
        {
            ListSalida.Items.Clear();
            ListPreservada.Items.Clear();
            ListEntrada.Items.Clear();
        }

        private void btnReporte_Click(object sender, EventArgs e)
        {
            Reporte reporte = new Reporte();
            reporte.Show();
            this.Close();
        }

        private void btnCargarArchivo_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string archivo = openFileDialog1.FileName;
            using (StreamReader fileReader = new StreamReader(archivo))
            {
                string stringReader;
                while ((stringReader = fileReader.ReadLine()) != null)
                {
                    ListEntrada.Items.Add(stringReader);
                }
            }
        }

        private void btnCompilar_Click_1(object sender, EventArgs e)
        {
            renglon = 0;
            while (renglon < ListEntrada.Items.Count)
            {
                ListEntrada.SelectedIndex = renglon;
                wlinea = ListEntrada.Text;
                BuscarTokens();
                renglon++;
            }
        }

        public string ObtenerUsuario(string usuario)
        {
            string IDUSARIO;
            string QUERY = "SELECT idusuario from Usuarios where Usuario=@usuario";
            conectar.Open();
            SqlCommand cmd = new SqlCommand(QUERY,conectar);
            cmd.Parameters.AddWithValue("@usuario", usuario);
            object result = cmd.ExecuteScalar();
            IDUSARIO = result.ToString();
            conectar.Close();
            return IDUSARIO;
        }

       

        //public void BuscarStrings()
        //{
        //    bool encontro = false;
        //    int renglon2 = 0;
        //    while (!encontro && renglon2 < ListString.Items.Count)
        //    {
        //        ListString.SelectedIndex = renglon2;
        //        if (string.Equals(token.ToUpper(), ListString.Text.ToUpper()))
        //        {
        //            encontro = true;
        //            direccion = renglon2;
        //        }
        //        renglon2++;
        //    }
        //    if (!encontro)
        //    {
        //        direccion = ListString.Items.Add(token);
        //        direccion = ListString.Items.Count - 1;
        //    }
        //}


    }
}