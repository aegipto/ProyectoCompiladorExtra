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

namespace ProyectoCompilador
{
    public partial class Reporte : Form
    {
        SqlConnection conexion = new SqlConnection("Data Source = DESKTOP-VCEO21F; Initial Catalog= SistemasProgramacion; Integrated Security = True ");
        public Reporte()
        {
            InitializeComponent();
            txtusuario.TextChanged += txtusuario_TextChanged;
            txtlenguaje.TextChanged += txtlenguaje_TextChanged;
        }

        private void Reporte_Load(object sender, EventArgs e)
        {
            string consulta = "select us.usuario 'Usuario', le.nombre 'Lenguaje', re.fecha_hora 'Fecha y Hora', re.nombre_archivo 'Archivo Salida' from Registro re INNER JOIN Usuarios us on re.idusuario = us.idusuario\tINNER JOIN Lenguajes le\ton re.idlenguaje = le.idlenguaje";
            SqlDataAdapter adapter = new SqlDataAdapter(consulta, conexion);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dgvReport.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dgvReport.Rows.Count > 0)
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Archivos de texto (*.txt)|*.txt";
                saveDialog.Title = "Guardar archivo de texto";
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveDialog.FileName;
                    try
                    {
                        using (StreamWriter writer = new StreamWriter(filePath))
                        {
                            // Escribir encabezados de columnas
                            for (int i = 0; i < dgvReport.Columns.Count; i++)
                            {
                                writer.Write(dgvReport.Columns[i].HeaderText);
                                if (i < dgvReport.Columns.Count - 1)
                                    writer.Write("\t"); // Tabulador como separador de columnas
                            }
                            writer.WriteLine();

                            // Escribir datos de celdas
                            foreach (DataGridViewRow row in dgvReport.Rows)
                            {
                                for (int i = 0; i < dgvReport.Columns.Count; i++)
                                {
                                    writer.Write(row.Cells[i].Value);
                                    if (i < dgvReport.Columns.Count - 1)
                                        writer.Write("\t"); // Tabulador como separador de columnas
                                }
                                writer.WriteLine();
                            }

                            writer.Close();
                        }

                        MessageBox.Show("Los datos se han exportado correctamente.", "Exportar a TXT", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Se produjo un error al exportar los datos: " + ex.Message, "Exportar a TXT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("No hay datos para exportar.", "Exportar a TXT", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        
    }

        private void btnCSV_Click(object sender, EventArgs e)
        {
            if (dgvReport.Rows.Count > 0)
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Archivo CSV (*.csv)|*.csv";
                saveDialog.Title = "Guardar archivo CSV";
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveDialog.FileName;

                    try
                    {
                        using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
                        {
                            // Escribir encabezados de columnas
                            for (int i = 0; i < dgvReport.Columns.Count; i++)
                            {
                                writer.Write(dgvReport.Columns[i].HeaderText);
                                if (i < dgvReport.Columns.Count - 1)
                                    writer.Write(","); // Coma como separador de columnas
                            }
                            writer.WriteLine();

                            // Escribir datos de celdas
                            foreach (DataGridViewRow row in dgvReport.Rows)
                            {
                                for (int i = 0; i < dgvReport.Columns.Count; i++)
                                {
                                    if (row.Cells[i].Value != null)
                                    {
                                        string value = row.Cells[i].Value.ToString();
                                        // Escapar las comas dentro del valor
                                        value = value.Replace(",", "\\,");
                                        writer.Write(value);
                                    }

                                    if (i < dgvReport.Columns.Count - 1)
                                        writer.Write(","); // Coma como separador de columnas
                                }
                                writer.WriteLine();
                            }

                            writer.Close();
                        }

                        MessageBox.Show("Los datos se han exportado correctamente.", "Exportar a CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Se produjo un error al exportar los datos: " + ex.Message, "Exportar a CSV", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("No hay datos para exportar.", "Exportar a CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        //private void ExportToExcel(DataGridView dataGridView, string filePath)
        //{
        //    // Crea un nuevo DataTable
        //    DataTable dt = new DataTable();

        //    // Recorre las columnas del DataGridView y agrega columnas al DataTable
        //    foreach (DataGridViewColumn column in dataGridView.Columns)
        //    {
        //        dt.Columns.Add(column.HeaderText, column.ValueType);
        //    }

        //    // Recorre las filas del DataGridView y agrega filas al DataTable
        //    foreach (DataGridViewRow row in dataGridView.Rows)
        //    {
        //        DataRow dataRow = dt.NewRow();
        //        foreach (DataGridViewCell cell in row.Cells)
        //        {
        //            dataRow[cell.ColumnIndex] = cell.Value;
        //        }
        //        dt.Rows.Add(dataRow);
        //    }

        //    // Exporta el DataTable al archivo XLSX utilizando la librería ExcelDataReader
        //    using (ExcelPackage excelPackage = new ExcelPackage())
        //    {
        //        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
        //        worksheet.Cells["A1"].LoadFromDataTable(dt, true);

        //        // Guarda el archivo XLSX en la ubicación especificada
        //        FileInfo excelFile = new FileInfo(filePath);
        //        excelPackage.SaveAs(excelFile);
        //    }
        //}

        private void btnXLSX_Click(object sender, EventArgs e)
        {
            string filePath = "C:\\Users\\alber\\OneDrive\\Escritorio\\Hola\\archivo.xlsx";
            ExportToExcel(dgvReport, filePath);
        }


        public void ExportToExcel(DataGridView dataGridView, string filePath)
        {
            // Crear una instancia de Excel
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet = workbook.ActiveSheet;

            try
            {
                // Encabezados de columna
                for (int i = 1; i <= dataGridView.Columns.Count; i++)
                {
                    sheet.Cells[1, i] = dataGridView.Columns[i - 1].HeaderText;
                }

                // Datos de las filas
                for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        sheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                    }
                }

                // Guardar el archivo Excel
                workbook.SaveAs(filePath);
                workbook.Close();
                excel.Quit();

                MessageBox.Show("Exportación exitosa");
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar a Excel: " + ex.Message);
            }
            finally
            {
                ReleaseObject(sheet);
                ReleaseObject(workbook);
                ReleaseObject(excel);
            }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Error al liberar el objeto: " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void txtusuario_TextChanged(object sender, EventArgs e)
        {
          string filtro = txtusuario.Text;
          (dgvReport.DataSource as DataTable).DefaultView.RowFilter = $"usuario LIKE '%{filtro}'";
        }

        private void txtusuario_KeyUp(object sender, KeyEventArgs e)
        {
            
        }

        private void txtlenguaje_TextChanged(object sender, EventArgs e)
        {
            string filtro = txtlenguaje.Text;
            (dgvReport.DataSource as DataTable).DefaultView.RowFilter = $"Lenguaje LIKE '%{filtro}'";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Form1 form = new Form1();
            //form.Show();
        }
    }
}
