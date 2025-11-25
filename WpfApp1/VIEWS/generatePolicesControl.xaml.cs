using Enla_C.DB;
using FirebirdSql.Data.FirebirdClient;
using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using Microsoft.Win32;
using Org.BouncyCastle.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Alignment = iText.Layout.Properties.HorizontalAlignment;
using Border = iText.Layout.Borders.Border;
using Image = iText.Layout.Element.Image;
using Paragraph = iText.Layout.Element.Paragraph;
using Table = iText.Layout.Element.Table;
using TextAlignment = iText.Layout.Properties.TextAlignment;

namespace Enla_C.VIEWS
{
    /// <summary>
    /// Lógica de interacción para UserControl1.xaml
    /// </summary>
    public partial class generatePolicesControl : UserControl
    {
        private static readonly Regex _regex = new Regex(@"^[0-9]*\.?[0-9]*$");

        public generatePolicesControl()
        {
            InitializeComponent();
        }

        private void radioSalidaMaquila_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (gridMateriaPrima != null)
                    gridMateriaPrima.Visibility = Visibility.Collapsed;
                if (gridMerma != null)
                    gridMerma.Visibility = Visibility.Collapsed;
                if (groupBoxCostosAdicionales != null)
                    groupBoxCostosAdicionales.Visibility = Visibility.Collapsed;

                CambiarAlmacen();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error inesperado:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void radioEntradaProcesado_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                gridMateriaPrima.Visibility = Visibility.Visible;
                gridMerma.Visibility = Visibility.Visible;
                groupBoxCostosAdicionales.Visibility = Visibility.Visible;

                CambiarAlmacen();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error inesperado:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void CambiarAlmacen()
        {
            ConectDB conect = ConectDB.Instance;

            if (radioSalidaMaquila.IsChecked == true)
            {
                if (comboAlmacenOrigen != null)
                    comboAlmacenOrigen.Text = conect.Alm1;
                if (comboAlmacenDestino != null)
                    comboAlmacenDestino.Text = conect.Alm2;
            }
            else if (radioEntradaProcesado.IsChecked == true)
            {
                comboAlmacenOrigen.Text = conect.Alm2;
                comboAlmacenDestino.Text = conect.Alm1;
            }
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            TextBox tb = sender as TextBox;
            string fullText = tb.Text.Insert(tb.CaretIndex, e.Text);
            e.Handled = !_regex.IsMatch(fullText);
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {

        }

        public void loadData()
        {
            try
            {
                datePickerFecha.SelectedDate = DateTime.Now.Date;
                datePickerFecha.DisplayDate = DateTime.Now.Date;

                ConectDB conect = ConectDB.Instance;

                comboAlmacenOrigen.Text = conect.Alm1;
                comboAlmacenDestino.Text = conect.Alm2;

                //DateTable PROV
                string tablesql = $"PROV{conect.NoEmpresa}";
                string sql = $"SELECT {tablesql}.CLAVE,{tablesql}.NOMBRE FROM {tablesql} ORDER BY {tablesql}.CLAVE";

                DataTable dataTable = conect.ExecuteQuery(sql);

                comboProveedor.ItemsSource = dataTable.DefaultView;
                comboProveedor.DisplayMemberPath = "NOMBRE";
                comboProveedor.SelectedValuePath = "CLAVE";

                //DateTable INVE
                tablesql = $"INVE{conect.NoEmpresa}";
                sql = $"SELECT {tablesql}.CVE_ART,{tablesql}.DESCR FROM {tablesql} ORDER BY {tablesql}.CVE_ART";

                dataTable = conect.ExecuteQuery(sql);

                comboProducto.ItemsSource = dataTable.DefaultView;
                comboProducto.DisplayMemberPath = "DESCR";
                comboProducto.SelectedValuePath = "CVE_ART";

                comboMateriaPrima.ItemsSource = dataTable.DefaultView;
                comboMateriaPrima.DisplayMemberPath = "DESCR";
                comboMateriaPrima.SelectedValuePath = "CVE_ART";

                comboMateriaPrima.IsEnabled = conect.EditaMP;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error inesperado:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void UnloadData()
        {
            datePickerFecha.SelectedDate = null;

            comboProveedor.SelectedIndex = -1;
            comboProveedor.Text = string.Empty;
            comboProveedor.ItemsSource = null;

            comboAlmacenOrigen.Text = string.Empty;
            comboAlmacenDestino.Text = string.Empty;

            comboProducto.SelectedIndex = -1;
            comboProducto.Text = string.Empty;
            comboProducto.ItemsSource = null;

            comboMateriaPrima.SelectedIndex = -1;
            comboMateriaPrima.Text = string.Empty;
            comboMateriaPrima.ItemsSource = null;

            textBoxCantidad.Text = string.Empty;
            textBoxMerma.Text = "0.0";
            textBoxAdicional1.Text = string.Empty;
            textBoxAdicional2.Text = string.Empty;
        }

        private void buttonAcceptGeneratePolicies_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(datePickerFecha.Text))
            {
                MessageBox.Show("El campo Fecha es obligatorio.", "Error de validación", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                datePickerFecha.Focus();
                return;
            }

            if (string.IsNullOrEmpty(comboProveedor.Text))
            {
                MessageBox.Show("El campo Proveedor es obligatorio.", "Error de validación", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                comboProveedor.Focus();
                return;
            }

            if (string.IsNullOrEmpty(comboAlmacenOrigen.Text))
            {
                MessageBox.Show("El campo Almacén Origen es obligatorio.", "Error de validación", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                comboAlmacenOrigen.Focus();
                return;
            }

            if (string.IsNullOrEmpty(comboAlmacenDestino.Text))
            {
                MessageBox.Show("El campo Almacén Destino es obligatorio.", "Error de validación", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                comboAlmacenDestino.Focus();
                return;
            }

            if (string.IsNullOrEmpty(comboProducto.Text))
            {
                MessageBox.Show("El campo Producto es obligatorio.", "Error de validación", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                comboProducto.Focus();
                return;
            }

            if (string.IsNullOrEmpty(textBoxCantidad.Text))
            {
                MessageBox.Show("El campo Cantidad es obligatorio.", "Error de validación", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                textBoxCantidad.Focus();
                return;
            }

            if (string.IsNullOrEmpty(comboMateriaPrima.Text) && (bool)radioEntradaProcesado.IsChecked)
            {
                MessageBox.Show("El campo Materia prima es obligatorio.", "Error de validación", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                comboMateriaPrima.Focus();
                return;
            }

            if (string.IsNullOrEmpty(textBoxMerma.Text) && (bool)radioEntradaProcesado.IsChecked)
            {
                MessageBox.Show("El campo Merma es obligatorio.", "Error de validación", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                textBoxMerma.Focus();
                return;
            }

            generatePolicesMov();
        }

        private void generatePolicesMov()
        {
            try
            {
                ConectDB conect = ConectDB.Instance;
                DateTime FECHA_MOV = (DateTime)datePickerFecha.SelectedDate;

                string CVE_ART_O = "";
                string CVE_ART_D = "";

                if ((bool)radioSalidaMaquila.IsChecked)
                {
                    CVE_ART_O = comboProducto.SelectedValue.ToString();
                    CVE_ART_D = comboProducto.SelectedValue.ToString();
                }
                else
                {
                    CVE_ART_O = comboMateriaPrima.SelectedValue.ToString();
                    CVE_ART_D = comboProducto.SelectedValue.ToString();
                }

                string prefix = radioSalidaMaquila.IsChecked == true ? "MQ" : "PP";
                string day = FECHA_MOV.ToString("dd");
                string month = FECHA_MOV.ToString("MM");
                string year = FECHA_MOV.ToString("yy");
                string timePart = DateTime.Now.ToString("HHmm");

                string CVE_DOC = $"{prefix}{day}{month}{year}-{timePart}";
                string CVE_CLPV = comboProveedor.SelectedValuePath;
                string DVK_BITA = $"BITA{conect.NoEmpresa}";
                string INVE = $"INVE{conect.NoEmpresa}";
                string MULT = $"MULT{conect.NoEmpresa}";
                string MINVE = $"MINVE{conect.NoEmpresa}";
                string TBLCONTROL = $"TBLCONTROL{conect.NoEmpresa}";
                string INVE_CLIB = $"INVE_CLIB{conect.NoEmpresa}";
                string DESCR_O;
                string UNI_MED_O;
                string DESCR_D;
                string COSTO_PROM_D;
                string UNI_MED_D;
                DateTime fechaStr = DateTime.Now;

                string sqlINVE = $"SELECT DESCR, EXIST, COSTO_PROM, UNI_MED, TIP_COSTEO FROM {INVE} WHERE CVE_ART = @cveArt";
                string sqlInveDescr = $@"SELECT DESCR FROM {INVE} WHERE CVE_ART = @cveArt";
                string sqlUltCve = $"SELECT ULT_CVE FROM {TBLCONTROL} WHERE ID_TABLA = @idTabla";
                string sqlSetUltCve = $@"UPDATE {TBLCONTROL} SET ULT_CVE = @numMov WHERE ID_TABLA = @idTabla";
                string sqlSetUltCveFolio = $@"UPDATE {TBLCONTROL} SET ULT_CVE = @cveFolio WHERE ID_TABLA = @idTabla";
                string sqlMult = $"SELECT EXIST FROM {MULT} WHERE CVE_ART = @cveArt AND CVE_ALM = @cveAlm";
                string sqlMultUpdateExistLess = $@"UPDATE {MULT} SET EXIST = EXIST - @cant WHERE CVE_ART = @cveArt AND CVE_ALM = @almacen";
                string sqlMultUpdateExistPlus = $@"UPDATE {MULT} SET EXIST = EXIST + @increment WHERE CVE_ART = @cveArt AND CVE_ALM = @almacen";
                string sqlMinve = $@"SELECT FIRST 1 COSTO_PROM_FIN FROM {MINVE} WHERE NUM_MOV = (SELECT MAX(NUM_MOV) FROM {MINVE} WHERE CVE_ART = @cveArt AND ALMACEN = @almacen) ORDER BY FECHAELAB DESC";
                string sqlMinveSum = $@"SELECT SUM((CANT * COSTO) * SIGNO) / SUM(CANT * SIGNO) AS COST FROM {MINVE} WHERE CVE_ART = @cveArt";
                string sqlMinveSum2 = $@"SELECT SUM(CANT * SIGNO) AS TotalCant, SUM((CANT * COSTO) * SIGNO) AS TotalCost FROM {MINVE} WHERE CVE_ART = @cveArt AND ALMACEN = @almacen";
                string sqlMinveCostoPromGral = $@"SELECT FIRST 1 COSTO_PROM_GRAL FROM {MINVE} WHERE NUM_MOV = (SELECT MAX(NUM_MOV) FROM {MINVE} WHERE CVE_ART = @cveArt) ORDER BY FECHAELAB DESC";
                string sqlInsertMinve = $@"INSERT INTO {MINVE} (CVE_ART, ALMACEN, NUM_MOV, CVE_CPTO, FECHA_DOCU, TIPO_DOC, REFER, CLAVE_CLPV, VEND, CANT, CANT_COST, PRECIO, COSTO, CVE_OBS, REG_SERIE, UNI_VENTA, E_LTPD, EXIST_G, EXISTENCIA, TIPO_PROD, FACTOR_CON, FECHAELAB, CVE_FOLIO, SIGNO, COSTEADO, COSTO_PROM_INI, COSTO_PROM_FIN, COSTO_PROM_GRAL, DESDE_INVE, MOV_ENLAZADO) VALUES (@cveArt, @almacen, @numMov, @cveCpto, @fechaDocu, @tipoDoc, @refer, @claveClpv, @vend, @cant, @cantCost, @precio, @costo, @cveObs, @regSerie, @uniVenta, @eLtpd, @existG, @existencia, @tipoProd, @factorCon, @fechaElab, @cveFolio, @signo, @costeado, @costoPromIni, @costoPromFin, @costoPromGral, @desdeInve, @movEnlazado)";
                string sqlInsertMinve2 = $@"INSERT INTO {MINVE} (CVE_ART, ALMACEN, NUM_MOV, CVE_CPTO, FECHA_DOCU, TIPO_DOC, REFER, CLAVE_CLPV, VEND, CANT, CANT_COST, PRECIO, COSTO, CVE_OBS, REG_SERIE, UNI_VENTA, E_LTPD, EXIST_G, EXISTENCIA, TIPO_PROD, FACTOR_CON, FECHAELAB, CVE_FOLIO, SIGNO, COSTEADO, COSTO_PROM_INI, COSTO_PROM_FIN, COSTO_PROM_GRAL, DESDE_INVE, MOV_ENLAZADO) VALUES (@cveArt, @almacen, @numMov, @cveCpto, @fechaMov, @tipoDoc, @refer, @claveClpv, @vend, @cant, @cantCost, @precio, @costo, @cveObs, @regSerie, @uniVenta, @eLtpd, @existG, @existencia, @tipoProd, @factorCon, @fechaElab, @cveFolio, @signo, @costeado, @costoPromIni, @costoPromFin, @costoPromGral, @desdeInve, @movEnlazado)";
                string sqlInsertDvk_bita = $@"INSERT INTO {DVK_BITA} (CVE_BITA, CVE_CLIE, CVE_CAMPANIA, CVE_ACTIVIDAD, FECHAHORA, CVE_USUARIO, OBSERVACIONES, STATUS, NOM_USUARIO) VALUES (@cveBita, @cveClie, @cveCampania, @cveActividad, @fechaHora, @cveUsuario, @observaciones, @status, @nomUsuario)";
                string sqlUpdateExistencia = $"UPDATE {INVE} SET EXIST = EXIST - @cant WHERE CVE_ART = @cveArt";
                string sqlUpdateExistencia2 = $"UPDATE {INVE} SET EXIST = EXIST + @cantidadNeta WHERE CVE_ART = @cveArt";
                string sqlUpdateCosto = $"UPDATE {INVE} SET ULT_COSTO = @nuevoCosto WHERE CVE_ART = @cveArt";
                string sqlUpdateCostoProm = $"UPDATE {INVE} SET COSTO_PROM = @costoPromedio WHERE CVE_ART = @cveArt";

                int NUM_ALM_O = int.Parse(comboAlmacenOrigen.Text);
                int NUM_ALM_D = int.Parse(comboAlmacenDestino.Text);
                int CANT = int.Parse(textBoxCantidad.Text);

                long NUM_MOV = 0;
                long NUM_MOV_ENLAZADO = 0;
                long CVE_FOLIO = 0;

                double EXIST_O = 0;
                double EXIST_D = 0;
                double COSTO_PROM_INI_O = 0;
                double COSTO_PROM_INI_D = 0;
                double COSTO_PROM_FIN_O = 0;
                double COSTO_PROM_FIN_D = 0;
                double COSTO_PROM_O = 0;
                double COSTO_PROM_GRAL_O = 0;
                double COSTO_PROM_GRAL_D = 0;
                double MERMA = 0.0;
                double VALOR_MERMA = 0.0;
                double EXIST_GRAL_O;
                double EXIST_GRAL_D;
                double COSTO_ADI1;
                double COSTO_ADI2;
                double COSTO_PROM;

                //cveArt CVE_ART_O
                DataTable dt = conect.ExecuteQuery(sqlINVE, new FbParameter("@cveArt", CVE_ART_O));

                if (dt.Rows.Count > 0)
                {
                    var row = dt.Rows[0];
                    DESCR_O = row["DESCR"]?.ToString();
                    EXIST_GRAL_O = Convert.ToDouble(row["EXIST"]);
                    UNI_MED_O = row["UNI_MED"]?.ToString();
                }
                else
                {
                    DESCR_O = string.Empty;
                    EXIST_GRAL_O = 0;
                    UNI_MED_O = string.Empty;
                }

                //idTabla 44
                dt = conect.ExecuteQuery(sqlUltCve, new FbParameter("@idTabla", 44));

                if (dt.Rows.Count > 0)
                {
                    var row = dt.Rows[0];
                    object val = row["ULT_CVE"];
                    if (val != DBNull.Value)
                        NUM_MOV = Convert.ToInt64(val);
                    else
                        NUM_MOV = 0;
                }
                else
                {
                    NUM_MOV = 0;
                }

                //idTabla 32
                dt = conect.ExecuteQuery(sqlUltCve, new FbParameter("@idTabla", 32));

                if (dt.Rows.Count > 0)
                {
                    var row = dt.Rows[0];
                    object val = row["ULT_CVE"];
                    if (val != DBNull.Value)
                        CVE_FOLIO = Convert.ToInt64(val) + 1;
                    else
                        CVE_FOLIO = 0;
                }
                else
                {
                    CVE_FOLIO = 0;
                }

                //cveArt CVE_ART_O cveAlm NUM_ALM_O
                dt = conect.ExecuteQuery(sqlMult, new FbParameter("@cveArt", CVE_ART_O), new FbParameter("@cveAlm", NUM_ALM_O));

                if (dt.Rows.Count > 0)
                {
                    var row = dt.Rows[0];
                    object val = row["EXIST"];
                    if (val != DBNull.Value)
                        EXIST_O = Convert.ToDouble(val);
                    else
                        EXIST_O = 0;
                }
                else
                {
                    EXIST_O = 0;
                }

                //cveArt CVE_ART_D cveAlm NUM_ALM_D
                dt = conect.ExecuteQuery(sqlMult, new FbParameter("@cveArt", CVE_ART_D), new FbParameter("@cveAlm", NUM_ALM_D));

                if (dt.Rows.Count > 0)
                {
                    var row = dt.Rows[0];
                    object val = row["EXIST"];
                    if (val != DBNull.Value)
                        EXIST_D = Convert.ToDouble(val);
                    else
                        EXIST_D = 0;
                }
                else
                {
                    EXIST_D = 0;
                }

                // Mostrar cuadro de mensaje en WPF
                if (EXIST_O < CANT)
                {
                    MessageBox.Show($"No hay existencias suficientes del producto:\r\n{CVE_ART_O}\r\n{EXIST_O:F2}", "Existencias insuficientes", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.FileName = CVE_DOC;

                    if (!(bool)saveFileDialog.ShowDialog())
                    {
                        MessageBox.Show("Proceso cancelado por el usuario.", "Información", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                    string rutaArchivo = saveFileDialog.FileName;

                    using (PdfWriter writer = new PdfWriter(rutaArchivo))
                    using (PdfDocument pdfDoc = new PdfDocument(writer))
                    using (Document document = new Document(pdfDoc, PageSize.LETTER))
                    {
                        PdfFont TipoLetra1 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                        PdfFont TipoLetra1Bold = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                        PdfFont TipoLetra2 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                        PdfFont TipoLetra3 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                        PdfFont TipoLetra4 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                        PdfFont TipoLetra5 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                        PdfFont TipoLetra6 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

                        float[] columnWidths = { 10f, 20f, 15f, 15f, 40f };
                        Table pdfPtable1 = new Table(2);
                        Table pdfPtable2 = new Table(2);
                        Table pdfPtable3 = new Table(columnWidths, true).SetBorderCollapse(BorderCollapsePropertyValue.SEPARATE);
                        Table pdfPtable4 = new Table(columnWidths, true);

                        pdfPtable1.SetPaddings(4f, 4f, 4f, 4f).SetWidth(UnitValue.CreatePercentValue(100f));
                        pdfPtable2.SetPaddings(4f, 4f, 4f, 4f).SetWidth(UnitValue.CreatePercentValue(95f));
                        pdfPtable3.SetPaddings(4f, 4f, 4f, 4f).SetWidth(UnitValue.CreatePercentValue(100f));
                        pdfPtable4.SetPaddings(4f, 4f, 4f, 4f).SetWidth(UnitValue.CreatePercentValue(100f));

                        pdfPtable2.SetMarginTop(10);
                        pdfPtable2.SetMarginBottom(10);

                        string titlePdf = "";
                        if ((bool)radioSalidaMaquila.IsChecked)
                            titlePdf = "SALIDA A MAQUILA";
                        else if ((bool)radioEntradaProcesado.IsChecked)
                            titlePdf = "ENTRADA DE PRODUCTO PROCESADO";

                        Cell cell1 = new Cell(1, 2)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(conect.RS)
                                .SetFont(TipoLetra1Bold)
                                .SetFontSize(18));
                        pdfPtable1.AddCell(cell1);

                        Cell cell2 = new Cell(1, 2)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(" ")
                                .SetFont(TipoLetra1)
                                .SetFontSize(10));
                        pdfPtable1.AddCell(cell2);

                        Cell cell3 = new Cell(1, 2)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(titlePdf)
                                .SetFont(TipoLetra1)
                                .SetFontSize(14));
                        pdfPtable1.AddCell(cell3);

                        Cell cell4 = new Cell(1, 2)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(" ")
                                .SetFont(TipoLetra1)
                                .SetFontSize(10));
                        pdfPtable1.AddCell(cell4);

                        Cell cell5 = new Cell(1, 2)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(" ")
                                .SetFont(TipoLetra1)
                                .SetFontSize(10));
                        pdfPtable1.AddCell(cell5);

                        Cell cell6 = new Cell(1, 4)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.LEFT)
                            .Add(new Paragraph($"PROVEEDOR: ({CVE_CLPV}) {comboProveedor.Text}")
                                .SetFont(TipoLetra2)
                                .SetFontSize(10));
                        pdfPtable2.AddCell(cell6);

                        Cell cell7 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.LEFT)
                            .Add(new Paragraph("FOLIO: " + CVE_DOC)
                                .SetFont(TipoLetra2)
                                .SetFontSize(10));
                        pdfPtable2.AddCell(cell7);

                        Cell cell8 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.RIGHT)
                            .Add(new Paragraph("FECHA: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm"))
                                .SetFont(TipoLetra2)
                                .SetFontSize(10));
                        pdfPtable2.AddCell(cell8);

                        Cell cell9 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.LEFT)
                            .Add(new Paragraph("ALMACÉN DE SALIDA: " + NUM_ALM_O.ToString())
                                .SetFont(TipoLetra2)
                                .SetFontSize(10));
                        pdfPtable2.AddCell(cell9);

                        Cell cell10 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.RIGHT)
                            .Add(new Paragraph("ALMACEN DE ENTRADA: " + NUM_ALM_D)
                                .SetFont(TipoLetra2)
                                .SetFontSize(10));
                        pdfPtable2.AddCell(cell10);

                        Border borderStyle = new SolidBorder(ColorConstants.BLACK, 0.5f);

                        Cell cell12 = new Cell(1, 1)
                            .SetBorder(borderStyle)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph("ALMACEN")
                                .SetFont(TipoLetra4)
                                .SetFontSize(10));
                        pdfPtable3.AddCell(cell12);

                        Cell cell13 = new Cell(1, 1)
                            .SetBorder(borderStyle)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph("CONCEPTO")
                                .SetFont(TipoLetra4)
                                .SetFontSize(10));
                        pdfPtable3.AddCell(cell13);

                        Cell cell14 = new Cell(1, 1)
                            .SetBorder(borderStyle)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph("CANTIDAD")
                                .SetFont(TipoLetra4)
                                .SetFontSize(10));
                        pdfPtable3.AddCell(cell14);

                        Cell cell15 = new Cell(1, 1)
                            .SetBorder(borderStyle)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph("CLAVE")
                                .SetFont(TipoLetra4)
                                .SetFontSize(10));
                        pdfPtable3.AddCell(cell15);

                        Cell cell16 = new Cell(1, 1)
                            .SetBorder(borderStyle)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph("DESCRIPCION")
                                .SetFont(TipoLetra4)
                                .SetFontSize(10));
                        pdfPtable3.AddCell(cell16);


                        COSTO_PROM_INI_O = 0.0;

                        try
                        {
                            dt = conect.ExecuteQuery(sqlMinve, new FbParameter("@cveArt", CVE_ART_O), new FbParameter("@almacen", NUM_ALM_O));

                            COSTO_PROM_INI_O = 0.0;
                            if (dt.Rows.Count > 0)
                            {
                                DataRow row = dt.Rows[0];
                                object valor = row["COSTO_PROM_FIN"];
                                if (valor != DBNull.Value)
                                    COSTO_PROM_INI_O = Convert.ToDouble(valor);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error al obtener COSTO_PROM_FIN: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        COSTO_PROM_O = COSTO_PROM_INI_O;
                        COSTO_PROM_FIN_O = COSTO_PROM_INI_O;
                        COSTO_PROM_GRAL_O = 0.0;

                        try
                        {
                            dt = conect.ExecuteQuery(sqlMinveSum, new FbParameter("@cveArt", CVE_ART_O));

                            if (dt.Rows.Count > 0)
                            {
                                DataRow row = dt.Rows[0];
                                object valor = row["COST"];
                                if (valor != DBNull.Value)
                                {
                                    COSTO_PROM_GRAL_O = Convert.ToDouble(valor);
                                }
                            }
                        }
                        catch (Exception)
                        {
                            //MessageBox.Show($"Error al obtener COSTO_PROM_GRAL_O: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            COSTO_PROM_GRAL_O = 0.0;
                        }

                        if (COSTO_PROM_GRAL_O == 0.0)
                        {
                            try
                            {
                                dt = conect.ExecuteQuery(sqlMinveCostoPromGral, new FbParameter("@cveArt", CVE_ART_O));
                                Console.WriteLine($"{dt.Rows[0]}");
                                if (dt.Rows.Count > 0)
                                {
                                    DataRow row = dt.Rows[0];
                                    object valor = row["COSTO_PROM_GRAL"];
                                    if (valor != DBNull.Value)
                                    {
                                        COSTO_PROM_GRAL_O = Convert.ToDouble(valor);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error al obtener COSTO_PROM_GRAL_O: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }

                        if ((bool)radioEntradaProcesado.IsChecked)
                        {
                            if (double.Parse(textBoxMerma.Text) != 0.0)
                            {
                                VALOR_MERMA = double.Parse(textBoxMerma.Text);
                                MERMA = VALOR_MERMA;
                            }
                            else
                                MERMA = 0.0;
                        }
                        else
                        {
                            MERMA = 0.0;
                        }

                        //// ISSUE: variable of a reference type
                        //long local1;
                        //// ISSUE: explicit reference operation
                        //long num2 = checked {long num2 = NUM_MOV + 1L;NUM_MOV = num2;}
                        //local1 = num2;
                        //// ISSUE: variable of a reference type
                        //long local2;
                        //// ISSUE: explicit reference operation
                        //long num3 = checked(^(local2 = ref this.CVE_FOLIO) + 1L);
                        //local2 = num3;
                        EXIST_GRAL_O -= CANT - MERMA;
                        EXIST_O -= CANT - MERMA;

                        try
                        {
                            conect.ExecuteNonQuery(sqlInsertMinve,
                        new FbParameter("@cveArt", CVE_ART_O),
                        new FbParameter("@almacen", NUM_ALM_O),
                        new FbParameter("@numMov", NUM_MOV),
                        new FbParameter("@cveCpto", 53),  // si CVE_CPTO tipo de operación
                        new FbParameter("@fechaDocu", FECHA_MOV),
                        new FbParameter("@tipoDoc", "M"),
                        new FbParameter("@refer", CVE_DOC),
                        new FbParameter("@claveClpv", CVE_CLPV),
                        new FbParameter("@vend", ""),
                        new FbParameter("@cant", CANT - MERMA),
                        new FbParameter("@cantCost", 0),
                        new FbParameter("@precio", 0),
                        new FbParameter("@costo", COSTO_PROM_O),
                        new FbParameter("@cveObs", 0),
                        new FbParameter("@regSerie", 0),
                        new FbParameter("@uniVenta", UNI_MED_O),
                        new FbParameter("@eLtpd", 0),
                        new FbParameter("@existG", EXIST_GRAL_O),
                        new FbParameter("@existencia", EXIST_O),
                        new FbParameter("@tipoProd", "P"),
                        new FbParameter("@factorCon", 1),
                        new FbParameter("@fechaElab", fechaStr),
                        new FbParameter("@cveFolio", CVE_FOLIO),
                        new FbParameter("@signo", -1),
                        new FbParameter("@costeado", "S"),
                        new FbParameter("@costoPromIni", COSTO_PROM_INI_O),
                        new FbParameter("@costoPromFin", COSTO_PROM_FIN_O),
                        new FbParameter("@costoPromGral", COSTO_PROM_GRAL_O),
                        new FbParameter("@desdeInve", "S"),
                        new FbParameter("@movEnlazado", 0));
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Verifique configuración de Base de datos\r\n\r\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }

                        string MENSAJE_BITACORA = $"Agregar movimientos {DVK_BITA} desde {NUM_MOV}";

                        try
                        {
                            conect.ExecuteNonQuery(sqlInsertDvk_bita,
                                new FbParameter("@cveBita", Convert.ToInt32(conect.ExecuteScalar($"SELECT COALESCE(MAX(CVE_BITA), 0) + 1 FROM {DVK_BITA}"))),
                                new FbParameter("@cveClie", CVE_CLPV),
                                new FbParameter("@cveCampania", 0),
                                new FbParameter("@cveActividad", 0),
                                new FbParameter("@fechaHora", DateTime.Now),
                                new FbParameter("@cveUsuario", 1),
                                new FbParameter("@observaciones", MENSAJE_BITACORA),
                                new FbParameter("@status", "F"),
                                new FbParameter("@nomUsuario", Environment.UserName));
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error al insertar en {DVK_BITA}:\r\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }

                        Cell cell17 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(NUM_ALM_O.ToString())
                                .SetFont(TipoLetra3)
                                .SetFontSize(10));
                        pdfPtable4.AddCell(cell17);

                        Cell cell18 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph("53 - Salida a fáb.")
                                .SetFont(TipoLetra3)
                                .SetFontSize(10));
                        pdfPtable4.AddCell(cell18);

                        Cell cell19 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(((CANT - MERMA) * -1.0).ToString("N2"))
                                .SetFont(TipoLetra3)
                                .SetFontSize(10));
                        pdfPtable4.AddCell(cell19);

                        Cell cell20 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(CVE_ART_O)
                                .SetFont(TipoLetra3)
                                .SetFontSize(10));
                        pdfPtable4.AddCell(cell20);

                        Cell cell21 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(DESCR_O)
                                .SetFont(TipoLetra3)
                                .SetFontSize(10));
                        pdfPtable4.AddCell(cell21);

                        if ((bool)radioEntradaProcesado.IsChecked & MERMA > 0.0)
                        {
                            EXIST_GRAL_O -= MERMA;
                            EXIST_O -= MERMA;

                            try
                            {
                                conect.ExecuteNonQuery(sqlInsertMinve,
                            new FbParameter("@cveArt", CVE_ART_O),
                            new FbParameter("@almacen", NUM_ALM_O),
                            new FbParameter("@numMov", NUM_MOV),
                            new FbParameter("@cveCpto", 55),  // según tu valor “CVE_CPTO = 55”
                            new FbParameter("@fechaDocu", FECHA_MOV),
                            new FbParameter("@tipoDoc", "M"),
                            new FbParameter("@refer", CVE_DOC),
                            new FbParameter("@claveClpv", CVE_CLPV),
                            new FbParameter("@vend", ""),  // según tu valor vacío
                            new FbParameter("@cant", MERMA),
                            new FbParameter("@cantCost", 0),
                            new FbParameter("@precio", 0),
                            new FbParameter("@costo", COSTO_PROM_O),
                            new FbParameter("@cveObs", 0),
                            new FbParameter("@regSerie", 0),
                            new FbParameter("@uniVenta", UNI_MED_O),
                            new FbParameter("@eLtpd", 0),
                            new FbParameter("@existG", EXIST_GRAL_O),
                            new FbParameter("@existencia", EXIST_O),
                            new FbParameter("@tipoProd", "P"),
                            new FbParameter("@factorCon", 1),
                            new FbParameter("@fechaElab", fechaStr),
                            new FbParameter("@cveFolio", CVE_FOLIO),
                            new FbParameter("@signo", -1),
                            new FbParameter("@costeado", "S"),
                            new FbParameter("@costoPromIni", COSTO_PROM_INI_O),
                            new FbParameter("@costoPromFin", COSTO_PROM_FIN_O),
                            new FbParameter("@costoPromGral", COSTO_PROM_GRAL_O),
                            new FbParameter("@desdeInve", "S"),
                            new FbParameter("@movEnlazado", 0));

                                conect.ExecuteNonQuery(sqlInsertDvk_bita,
                                    new FbParameter("@cveBita", Convert.ToInt32(conect.ExecuteScalar($"SELECT COALESCE(MAX(CVE_BITA), 0) + 1 FROM {DVK_BITA}"))),
                                    new FbParameter("@cveClie", CVE_CLPV),
                                    new FbParameter("@cveCampania", 0),
                                    new FbParameter("@cveActividad", 0),
                                    new FbParameter("@fechaHora", DateTime.Now),
                                    new FbParameter("@cveUsuario", Environment.UserName),
                                    new FbParameter("@observaciones", MENSAJE_BITACORA),
                                    new FbParameter("@status", "F"),
                                    new FbParameter("@nomUsuario", Environment.UserName));
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Verifique configuración de Base de datos\r\n\r\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            }

                            Cell cell22 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(NUM_ALM_O.ToString())
                            .SetFont(TipoLetra3)
                            .SetFontSize(10));
                            pdfPtable4.AddCell(cell22);

                            Cell cell23 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph("55 - Mermas"))
                                .SetFont(TipoLetra3)
                                .SetFontSize(10);
                            pdfPtable4.AddCell(cell23);

                            Cell cell24 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph((MERMA * -1.0).ToString("N2")))
                                .SetFont(TipoLetra3)
                                .SetFontSize(10);
                            pdfPtable4.AddCell(cell24);

                            Cell cell25 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph(CVE_ART_O))
                                .SetFont(TipoLetra3)
                                .SetFontSize(10);
                            pdfPtable4.AddCell(cell25);

                            Cell cell26 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph(DESCR_O))
                                .SetFont(TipoLetra3)
                                .SetFontSize(10);
                            pdfPtable4.AddCell(cell26);
                        }

                        if ((bool)radioEntradaProcesado.IsChecked)
                        {
                            try
                            {
                                dt = conect.ExecuteQuery(sqlINVE, new FbParameter("@cveArt", CVE_ART_D));

                                if (dt.Rows.Count > 0)
                                {
                                    DataRow row = dt.Rows[0];
                                    DESCR_D = row["DESCR"]?.ToString();
                                    EXIST_GRAL_D = row["EXIST"] != DBNull.Value ? Convert.ToDouble(row["EXIST"]) : 0.0;
                                    COSTO_PROM_D = row["COSTO_PROM"]?.ToString();
                                    UNI_MED_D = row["UNI_MED"]?.ToString();
                                }
                                else
                                {
                                    DESCR_D = string.Empty;
                                    EXIST_GRAL_D = 0.0;
                                    COSTO_PROM_D = string.Empty;
                                    UNI_MED_D = string.Empty;
                                }

                                if (!string.IsNullOrEmpty(textBoxAdicional1.Text))
                                    COSTO_ADI1 = Convert.ToDouble(textBoxAdicional1.Text);
                                else
                                    COSTO_ADI1 = 0.0;

                                if (!string.IsNullOrEmpty(textBoxAdicional2.Text))
                                    COSTO_ADI2 = Convert.ToDouble(textBoxAdicional2.Text);
                                else
                                    COSTO_ADI2 = 0.0;

                                double baseCantidad = CANT - MERMA;
                                if (baseCantidad > 0)
                                {
                                    if (COSTO_ADI1 > 0.0)
                                        COSTO_ADI1 /= baseCantidad;
                                    if (COSTO_ADI2 > 0.0)
                                        COSTO_ADI2 /= baseCantidad;
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Verifique configuración de Base de datos\r\n\r\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }
                        else
                        {
                            try
                            {
                                dt = conect.ExecuteQuery(sqlInveDescr, new FbParameter("@cveArt", CVE_ART_D));

                                if (dt.Rows.Count > 0)
                                {
                                    DataRow row = dt.Rows[0];
                                    DESCR_D = row["DESCR"]?.ToString();
                                }
                                else
                                {
                                    DESCR_D = string.Empty;
                                }

                                COSTO_PROM_D = COSTO_PROM_O.ToString();
                                UNI_MED_D = UNI_MED_O;
                                COSTO_ADI1 = 0.0;
                                COSTO_ADI2 = 0.0;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Verifique configuración de Base de datos\r\n\r\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }

                        NUM_MOV_ENLAZADO = NUM_MOV;
                        EXIST_GRAL_D = !(bool)radioEntradaProcesado.IsChecked ? EXIST_GRAL_O + (CANT - MERMA) : EXIST_D + (CANT - MERMA);
                        EXIST_D += CANT - MERMA;
                        COSTO_PROM_INI_D = 0.0;

                        try
                        {
                            dt = conect.ExecuteQuery(sqlMinve, new FbParameter("@cveArt", CVE_ART_D), new FbParameter("@almacen", NUM_ALM_D));

                            if (dt.Rows.Count > 0)
                            {
                                object val = dt.Rows[0]["COSTO_PROM_FIN"];
                                if (val != DBNull.Value)
                                    COSTO_PROM_INI_D = Convert.ToDouble(val);
                            }

                            COSTO_PROM_FIN_D = 0.0;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error al obtener COSTO_PROM_FIN para D: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }


                        try
                        {
                            dt = conect.ExecuteQuery(sqlMinveSum2, new FbParameter("@cveArt", CVE_ART_D), new FbParameter("@almacen", NUM_ALM_D));

                            double totalCant = 0.0;
                            double totalCost = 0.0;

                            if (dt.Rows.Count > 0)
                            {
                                DataRow row = dt.Rows[0];
                                object cantObj = row["TotalCant"];
                                object costObj = row["TotalCost"];

                                if (cantObj != DBNull.Value)
                                    totalCant = Convert.ToDouble(cantObj);

                                if (costObj != DBNull.Value)
                                    totalCost = Convert.ToDouble(costObj);
                            }

                            COSTO_PROM_FIN_D = (totalCost + CANT * (COSTO_PROM_O + COSTO_ADI1 + COSTO_ADI2)) / (totalCant + CANT);
                            COSTO_PROM_GRAL_D = 0.0;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error al calcular costo prom final D: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        try
                        {
                            dt = conect.ExecuteQuery(sqlMinveSum, new FbParameter("@cveArt", CVE_ART_D));

                            if (dt.Rows.Count > 0)
                            {
                                object val = dt.Rows[0]["COST"];
                                if (val != DBNull.Value)
                                {
                                    COSTO_PROM_GRAL_D = Convert.ToDouble(val);
                                }
                                else
                                {
                                    COSTO_PROM_GRAL_D = 0.0;
                                }
                            }
                            else
                            {
                                COSTO_PROM_GRAL_D = 0.0;
                            }
                        }
                        catch (Exception)
                        {
                            //MessageBox.Show($"Error al obtener COSTO_PROM_GRAL_D: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            COSTO_PROM_GRAL_D = 0.0;
                        }

                        if (COSTO_PROM_GRAL_D == 0.0)
                        {
                            try
                            {
                                dt = conect.ExecuteQuery(sqlMinveCostoPromGral, new FbParameter("@cveArt", CVE_ART_D));

                                if (dt.Rows.Count > 0)
                                {
                                    object val = dt.Rows[0]["COSTO_PROM_GRAL"];
                                    if (val != DBNull.Value)
                                    {
                                        COSTO_PROM_GRAL_D = Convert.ToDouble(val);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error al obtener COSTO_PROM_GRAL_D: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }

                        if (CVE_ART_O != CVE_ART_D)
                        {
                            NUM_MOV_ENLAZADO = 0L;
                        }

                        try
                        {
                            conect.ExecuteNonQuery(sqlInsertMinve2,
                        new FbParameter("@cveArt", CVE_ART_D),
                        new FbParameter("@almacen", NUM_ALM_D),
                        new FbParameter("@numMov", NUM_MOV),
                        new FbParameter("@cveCpto", 3),
                        new FbParameter("@fechaMov", FECHA_MOV),
                        new FbParameter("@tipoDoc", "M"),
                        new FbParameter("@refer", CVE_DOC),
                        new FbParameter("@claveClpv", CVE_CLPV),
                        new FbParameter("@vend", ""),
                        new FbParameter("@cant", CANT - MERMA),
                        new FbParameter("@cantCost", 0),
                        new FbParameter("@precio", 0),
                        // Suma de costos como en tu lógica original
                        new FbParameter("@costo", COSTO_PROM_O + COSTO_ADI1 + COSTO_ADI2),
                        new FbParameter("@cveObs", 0),
                        new FbParameter("@regSerie", 0),
                        new FbParameter("@uniVenta", UNI_MED_D),
                        new FbParameter("@eLtpd", 0),
                        new FbParameter("@existG", EXIST_GRAL_D),
                        new FbParameter("@existencia", EXIST_D),
                        new FbParameter("@tipoProd", "P"),
                        new FbParameter("@factorCon", 1),
                        new FbParameter("@fechaElab", fechaStr),
                        new FbParameter("@cveFolio", CVE_FOLIO),
                        new FbParameter("@signo", 1),
                        new FbParameter("@costeado", "S"),
                        new FbParameter("@costoPromIni", COSTO_PROM_INI_D),
                        new FbParameter("@costoPromFin", COSTO_PROM_FIN_D),
                        new FbParameter("@costoPromGral", COSTO_PROM_GRAL_D),
                        new FbParameter("@desdeInve", "S"),
                        new FbParameter("@movEnlazado", NUM_MOV_ENLAZADO));
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Verifique configuración de Base de datos\r\n\r\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        Cell cell27 = new Cell(1, 1)
                            .SetBorder(Border.NO_BORDER)
                            .SetTextAlignment(TextAlignment.CENTER)
                            .Add(new Paragraph(NUM_ALM_D.ToString()))
                            .SetFont(TipoLetra3)
                            .SetFontSize(10);
                        pdfPtable4.AddCell(cell27);

                        Cell cell28 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph("3 - Entrada de fáb."))
                                .SetFont(TipoLetra3)
                                .SetFontSize(10);
                        pdfPtable4.AddCell(cell28);

                        Cell cell29 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph((CANT - MERMA).ToString("N2")))
                                .SetFont(TipoLetra3)
                                .SetFontSize(10);
                        pdfPtable4.AddCell(cell29);

                        Cell cell30 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph(CVE_ART_D))
                                .SetFont(TipoLetra3)
                                .SetFontSize(10);
                        pdfPtable4.AddCell(cell30);

                        Cell cell31 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph(DESCR_D))
                                .SetFont(TipoLetra3)
                                .SetFontSize(10);
                        pdfPtable4.AddCell(cell31);

                        try
                        {
                            conect.ExecuteNonQuery(sqlMultUpdateExistLess,
                        new FbParameter("@cant", CANT),
                        new FbParameter("@cveArt", CVE_ART_O),
                        new FbParameter("@almacen", NUM_ALM_O));

                            conect.ExecuteNonQuery(sqlMultUpdateExistPlus,
                                new FbParameter("@increment", CANT - MERMA),
                                new FbParameter("@cveArt", CVE_ART_D),
                                new FbParameter("@almacen", NUM_ALM_D));

                            conect.ExecuteNonQuery(sqlSetUltCve,
                                new FbParameter("@numMov", NUM_MOV),
                                new FbParameter("@idTabla", 44));

                            conect.ExecuteNonQuery(sqlSetUltCveFolio,
                                new FbParameter("@cveFolio", CVE_FOLIO),
                                new FbParameter("@idTabla", 32));
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Verifique configuración de Base de datos\r\n\r\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        if ((bool)radioEntradaProcesado.IsChecked)
                        {
                            try
                            {
                                conect.ExecuteNonQuery(sqlUpdateExistencia,
                            new FbParameter("@cant", CANT),
                            new FbParameter("@cveArt", CVE_ART_O));

                                object cantidadNeta = CANT - MERMA;
                                conect.ExecuteNonQuery(sqlUpdateExistencia2,
                                    new FbParameter("@cantidadNeta", cantidadNeta),
                                    new FbParameter("@cveArt", CVE_ART_D));

                                double nuevoCosto = double.Parse(COSTO_PROM_D);
                                conect.ExecuteNonQuery(sqlUpdateCosto,
                                    new FbParameter("@nuevoCosto", nuevoCosto),
                                    new FbParameter("@cveArt", CVE_ART_D));

                                COSTO_PROM = nuevoCosto;
                                conect.ExecuteNonQuery(sqlUpdateCostoProm,
                                    new FbParameter("@costoPromedio", COSTO_PROM_FIN_D),
                                    new FbParameter("@cveArt", CVE_ART_D));
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Verifique configuración de Base de datos\r\n\r\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }

                        try
                        {
                            document.SetMargins(30f, 30f, 30f, 30f);

                            string logoPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.jpg");
                            ImageData data = ImageDataFactory.Create(logoPath);
                            Image instance = new Image(data);
                            instance.SetFixedPosition(40f, 700f);
                            instance.SetHorizontalAlignment(Alignment.LEFT);

                            if (instance != null)
                                document.Add(instance);

                            document.Add(pdfPtable1);
                            document.Add(pdfPtable2);
                            document.Add(pdfPtable3);
                            document.Add(pdfPtable4);

                            Process.Start(rutaArchivo);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("No se pudo generar el documento PDF\r\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Verifique configuración de Base de datos\r\n\r\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
