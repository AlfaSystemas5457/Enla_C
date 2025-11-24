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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
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
    /// Lógica de interacción para cancelMovs.xaml
    /// </summary>
    public partial class cancelMovement : UserControl
    {
        public cancelMovement()
        {
            InitializeComponent();
        }

        private void ButtonAccept_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(textBoxFolio.Text))
            {
                MessageBox.Show("El campo Folio es obligatorio.", "Error de validación", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                textBoxFolio.Focus();
                return;
            }

            CancelMovement();
        }

        public void CancelMovement()
        {
            try
            {
                ConectDB conect = ConectDB.Instance;

                DateTime F_MOV = DateTime.Now;
                DateTime FECHA_MOV = F_MOV;
                DateTime FECHA_IMP = F_MOV;
                DateTime FECHA_ELAB = F_MOV;
                DateTime FECHA_PREFIX = F_MOV;

                string DVK_BITA = $"BITA{conect.NoEmpresa}";
                string INVE = $"INVE{conect.NoEmpresa}";
                string MULT = $"MULT{conect.NoEmpresa}";
                string MINVE = $"MINVE{conect.NoEmpresa}";
                string TBLCONTROL = $"TBLCONTROL{conect.NoEmpresa}";
                string INVE_CLIB = $"INVE_CLIB{conect.NoEmpresa}";
                string PROV = $"PROV{conect.NoEmpresa}";
                string CONM = $"CONM{conect.NoEmpresa}";
                string DESCR_O;
                string CVE_DOC = textBoxFolio.Text;
                string prefijoREFER = CVE_DOC.Length >= 2 ? CVE_DOC.Substring(0, 2) : "";
                string REFER = $"{(prefijoREFER.Equals("MQ", StringComparison.OrdinalIgnoreCase) ? "MQC" : "PPC")}{FECHA_PREFIX.ToString("ddMMyy-HHmm")}";
                string CVE_CLPV = "";
                string NOMBRE = "";
                string ERR_EXIST;
                string CVE_ART;

                string sqlLeerDetalleMovimientoOriginal = $@"SELECT CVE_ART, ALMACEN, CVE_CPTO, CANT, CLAVE_CLPV FROM {MINVE} WHERE REFER = @cveDoc AND TIPO_DOC = 'M'";
                string sqlLeerExistenciaAlmacen = $@"SELECT EXIST FROM {MULT} WHERE CVE_ART = @cveArt AND CVE_ALM = @numAlm";
                string sqlLeerDetallePDFConProveedor = $@"SELECT T1.CVE_ART, T1.ALMACEN, T1.CVE_CPTO, T1.CLAVE_CLPV, T2.NOMBRE, T1.CANT, T1.COSTO, T1.UNI_VENTA FROM {MINVE} T1 LEFT JOIN {PROV} T2 ON T2.CLAVE = T1.CLAVE_CLPV WHERE T1.REFER = @cveDoc AND T1.TIPO_DOC = 'M'";
                string sqlLeerDatosGeneralesArticulo = $@"SELECT DESCR, EXIST, COSTO_PROM, UNI_MED FROM {INVE} WHERE CVE_ART = @cveArt";
                string sqlLeerUltimaClaveMovimiento = $@"SELECT ULT_CVE FROM {TBLCONTROL} WHERE ID_TABLA = 44";
                string sqlLeerUltimoFolio = $@"SELECT ULT_CVE FROM {TBLCONTROL} WHERE ID_TABLA = 32";
                string sqlLeerExistenciaAlmacenCheck = $@"SELECT EXIST FROM {MULT} WHERE CVE_ART = @cveArt AND CVE_ALM = @numAlm";
                string sqlLeerDescripcionConcepto = $@"SELECT DESCR FROM {CONM} WHERE CVE_CPTO = @cveCptoRev";
                string sqlInsertarMovimientoCancelacion = $@"INSERT INTO {MINVE} (CVE_ART, ALMACEN, NUM_MOV, CVE_CPTO, FECHA_DOCU, TIPO_DOC, REFER, CLAVE_CLPV, VEND, CANT, CANT_COST, PRECIO, COSTO, CVE_OBS, REG_SERIE, UNI_VENTA, E_LTPD, EXIST_G, EXISTENCIA, TIPO_PROD, FACTOR_CON, FECHAELAB, CVE_FOLIO, SIGNO, COSTEADO, COSTO_PROM_INI, COSTO_PROM_FIN, COSTO_PROM_GRAL, DESDE_INVE, MOV_ENLAZADO) VALUES (@cveArt, @numAlm, @numMov, @cveCptoRev, @fechaMov, 'M', @refer, @cveClpv, '', @cant, 0, 0, @costo, 0, 0, @uniMed, 0, @existGral, @exist, 'P', 1, @fechaElab, @cveFolio, @signo, 'S', @costoPromIni, @costoPromFin, @costoPromGral, 'S', 0)";
                string sqlInsertarBitacoraCancelacion = $@"INSERT INTO {DVK_BITA} (CVE_BITA, CVE_CLIE, CVE_CAMPANIA, CVE_ACTIVIDAD, FECHAHORA, CVE_USUARIO, OBSERVACIONES, STATUS, NOM_USUARIO) VALUES (@cveBita, @cveClie, @cveCampania, @cveActividad, @fechaHora, @cveUsuario, @observaciones, @status, @nomUsuario)";
                string sqlActualizarExistenciaAlmacen = $@"UPDATE {MULT} SET EXIST = @existNueva WHERE CVE_ART = @cveArt AND CVE_ALM = @numAlm";
                string sqlActualizarExistenciaGeneral = $@"UPDATE {INVE} SET EXIST = @existGralNueva WHERE CVE_ART = @cveArt";
                string sqlActualizarUltimaClaveMovimiento = $@"UPDATE {TBLCONTROL} SET ULT_CVE = @numMov WHERE ID_TABLA = 44";
                string sqlActualizarUltimoFolio = $@"UPDATE {TBLCONTROL} SET ULT_CVE = @cveFolio WHERE ID_TABLA = 32";

                double NUM_ALM;
                double CVE_CPTO;
                double CANT;
                double REGS;
                double EXIST;

                try
                {
                    DataTable dtDetalle1 = conect.ExecuteQuery(sqlLeerDetalleMovimientoOriginal, new FbParameter("@cveDoc", CVE_DOC));

                    REGS = 0;
                    ERR_EXIST = "";

                    foreach (DataRow row in dtDetalle1.Rows)
                    {
                        CVE_ART = row["CVE_ART"].ToString();
                        NUM_ALM = Convert.ToInt32(row["ALMACEN"]);
                        CVE_CPTO = Convert.ToInt32(row["CVE_CPTO"]);
                        CANT = Convert.ToDouble(row["CANT"]);
                        CVE_CLPV = row["CLAVE_CLPV"].ToString();

                        if (CVE_CPTO < 51)
                        {
                            EXIST = 0.0;

                            object result = conect.ExecuteScalar(sqlLeerExistenciaAlmacen,
                                new FbParameter("@cveArt", CVE_ART),
                                new FbParameter("@numAlm", NUM_ALM));

                            if (result != null && result != DBNull.Value)
                            {
                                EXIST = Convert.ToDouble(result);
                            }

                            if (EXIST < CANT)
                            {
                                ERR_EXIST += $"{CVE_ART} existencias: {EXIST}, requeridas: {CANT}\r\n";
                            }
                        }

                        REGS++;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Verifique configuración de Base de datos (Consulta Detalle)\r\n\r\n{ex.Message}", "Error de BD", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (REGS == 0)
                {
                    MessageBox.Show("No se encontró el documento indicado.\r\nVerifique e intente de nuevo.", "Proceso cancelado. 🤚", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!string.IsNullOrEmpty(ERR_EXIST))
                {
                    MessageBox.Show($"No hay existencias suficientes de los siguientes productos:\r\n\r\n{ERR_EXIST}\r\n\r\nEl proceso será cancelado.", "Proceso cancelado. 🛑", MessageBoxButton.OK, MessageBoxImage.Stop);
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Archivos PDF (*.pdf)|*.pdf|Todos los archivos (*.*)|*.*";
                saveFileDialog.FilterIndex = 1;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = REFER;

                if (!(bool)saveFileDialog.ShowDialog())
                {
                    MessageBox.Show("Proceso cancelado por el usuario.", "Información", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                string rutaPDF = saveFileDialog.FileName;

                using (PdfWriter writer = new PdfWriter(rutaPDF))
                using (PdfDocument pdfDoc = new PdfDocument(writer))
                using (Document document = new Document(pdfDoc, PageSize.LETTER))
                {
                    PdfFont TipoLetra1 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                    PdfFont TipoLetra1Bold = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    PdfFont TipoLetra2 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                    PdfFont TipoLetra3 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                    PdfFont TipoLetra4 = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

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

                    string logoPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.jpg");
                    ImageData data = ImageDataFactory.Create(logoPath);
                    Image instance = new Image(data);
                    instance.SetFixedPosition(40f, 700f);
                    instance.SetHorizontalAlignment(Alignment.LEFT);

                    string prefijoTitulo = CVE_DOC.Length >= 2 ? CVE_DOC.Substring(0, 2) : "";
                    string tituloDoc = prefijoTitulo.Equals("MQ", StringComparison.OrdinalIgnoreCase) ? "CANCELACION SALIDA A MAQUILA" : "CANCELACION ENTRADA DE PRODUCTO PROCESADO";

                    // Celda 1: Razón Social (this.RS)
                    Cell cell1 = new Cell(1, 2)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph(conect.RS)
                                    .SetFont(TipoLetra1Bold)
                                    .SetFontSize(18));
                    pdfPtable1.AddCell(cell1);

                    // Celda 2: Espacio en blanco
                    Cell cell2 = new Cell(1, 2)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph(" ")
                                    .SetFont(TipoLetra1)
                                    .SetFontSize(10));
                    pdfPtable1.AddCell(cell2);

                    // Celda 3: Tipo de Cancelación (str9)
                    Cell cell3 = new Cell(1, 2)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph(tituloDoc)
                                    .SetFont(TipoLetra1)
                                    .SetFontSize(14));
                    pdfPtable1.AddCell(cell3);

                    // Celda 4: Espacio en blanco
                    Cell cell4 = new Cell(1, 2)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph(" ")
                                    .SetFont(TipoLetra1)
                                    .SetFontSize(10));
                    pdfPtable1.AddCell(cell4);

                    // Celda 5: Espacio en blanco
                    Cell cell5 = new Cell(1, 2)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph(" ")
                                    .SetFont(TipoLetra1)
                                    .SetFontSize(10));
                    pdfPtable1.AddCell(cell5);

                    pdfPtable2.SetMarginTop(10);
                    pdfPtable2.SetMarginBottom(10);

                    // Celda 6: Proveedor
                    Cell cell6 = new Cell(1, 4)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.LEFT)
                                .Add(new Paragraph($"PROVEEDOR: ({CVE_CLPV}) {NOMBRE}")
                                    .SetFont(TipoLetra2)
                                    .SetFontSize(10));
                    pdfPtable2.AddCell(cell6);

                    // Celda 7: Folio de Cancelación (this.REFER)
                    Cell cell7 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.LEFT)
                                .Add(new Paragraph("FOLIO: " + REFER)
                                    .SetFont(TipoLetra2)
                                    .SetFontSize(10));
                    pdfPtable2.AddCell(cell7);

                    // Celda 8: Fecha de Impresión (this.FECHA_IMP)
                    Cell cell8 = new Cell(1, 1)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .Add(new Paragraph("FECHA: " + FECHA_IMP)
                                    .SetFont(TipoLetra2)
                                    .SetFontSize(10));
                    pdfPtable2.AddCell(cell8);

                    // Celda 9: Folio Cancelado (this.CVE_DOC)
                    Cell cell9 = new Cell(1, 4)
                                .SetBorder(Border.NO_BORDER)
                                .SetTextAlignment(TextAlignment.LEFT)
                                .Add(new Paragraph("FOLIO CANCELADO: " + CVE_DOC)
                                    .SetFont(TipoLetra2)
                                    .SetFontSize(10));
                    pdfPtable2.AddCell(cell9);

                    Border borderStyle = new SolidBorder(ColorConstants.BLACK, 0.5f);

                    // Celda 11: ALMACEN
                    Cell cell11 = new Cell(1, 1)
                                .SetBorder(borderStyle)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph("ALMACEN")
                                    .SetFont(TipoLetra4)
                                    .SetFontSize(10));
                    pdfPtable3.AddCell(cell11);

                    // Celda 12: CONCEPTO
                    Cell cell12 = new Cell(1, 1)
                                .SetBorder(borderStyle)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph("CONCEPTO")
                                    .SetFont(TipoLetra4)
                                    .SetFontSize(10));
                    pdfPtable3.AddCell(cell12);

                    // Celda 13: CANTIDAD
                    Cell cell13 = new Cell(1, 1)
                                .SetBorder(borderStyle)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph("CANTIDAD")
                                    .SetFont(TipoLetra4)
                                    .SetFontSize(10));
                    pdfPtable3.AddCell(cell13);

                    // Celda 14: CLAVE
                    Cell cell14 = new Cell(1, 1)
                                .SetBorder(borderStyle)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph("CLAVE")
                                    .SetFont(TipoLetra4)
                                    .SetFontSize(10));
                    pdfPtable3.AddCell(cell14);

                    // Celda 15: DESCRIPCION
                    Cell cell15 = new Cell(1, 1)
                                .SetBorder(borderStyle)
                                .SetTextAlignment(TextAlignment.CENTER)
                                .Add(new Paragraph("DESCRIPCION")
                                    .SetFont(TipoLetra4)
                                    .SetFontSize(10));
                    pdfPtable3.AddCell(cell15);

                    try
                    {
                        long NUM_MOV = Convert.ToInt64(conect.ExecuteScalar(sqlLeerUltimaClaveMovimiento)) + 1;
                        long CVE_FOLIO = Convert.ToInt64(conect.ExecuteScalar(sqlLeerUltimoFolio)) + 1;

                        DataTable dtDetalle = conect.ExecuteQuery(sqlLeerDetallePDFConProveedor, new FbParameter("@cveDoc", CVE_DOC));
                        foreach (DataRow row in dtDetalle.Rows)
                        {
                            CVE_ART = row["CVE_ART"].ToString();
                            NUM_ALM = Convert.ToDouble(row["ALMACEN"]);
                            CVE_CPTO = Convert.ToDouble(row["CVE_CPTO"]);
                            CVE_CLPV = row["CLAVE_CLPV"].ToString();
                            NOMBRE = row["NOMBRE"].ToString();
                            CANT = Convert.ToDouble(row["CANT"]);
                            string UNI_MED = row["UNI_VENTA"].ToString();

                            DataTable dtArt = conect.ExecuteQuery(sqlLeerDatosGeneralesArticulo, new FbParameter("@cveArt", CVE_ART));

                            double EXIST_GRAL = Convert.ToDouble(dtArt.Rows[0]["EXIST"]);
                            double COSTO_PROM = Convert.ToDouble(dtArt.Rows[0]["COSTO_PROM"]);
                            string DESCR = dtArt.Rows[0]["DESCR"].ToString();

                            object exAlm = conect.ExecuteScalar(sqlLeerExistenciaAlmacen,
                                new FbParameter("@cveArt", CVE_ART),
                                new FbParameter("@numAlm", NUM_ALM));

                            EXIST = exAlm == null || exAlm == DBNull.Value ? 0 : Convert.ToDouble(exAlm);

                            int CVE_CPTO_REV = 0;
                            int SIGNO = 0;

                            if (CVE_CPTO >= 51)
                            {
                                SIGNO = 1;
                                EXIST += CANT;
                                EXIST_GRAL += CANT;
                                if (CVE_CPTO == 53) CVE_CPTO_REV = 10;
                                if (CVE_CPTO == 55) CVE_CPTO_REV = 11;
                            }
                            else
                            {
                                SIGNO = -1;
                                EXIST -= CANT;
                                EXIST_GRAL -= CANT;
                                if (CVE_CPTO == 3) CVE_CPTO_REV = 60;
                            }

                            conect.ExecuteNonQuery(sqlInsertarMovimientoCancelacion,
                                new FbParameter("@cveArt", CVE_ART),
                                new FbParameter("@numAlm", NUM_ALM),
                                new FbParameter("@numMov", NUM_MOV),
                                new FbParameter("@cveCptoRev", CVE_CPTO_REV),
                                new FbParameter("@fechaMov", FECHA_MOV),
                                new FbParameter("@refer", REFER),
                                new FbParameter("@cveClpv", CVE_CLPV),
                                new FbParameter("@cant", CANT),
                                new FbParameter("@costo", COSTO_PROM),
                                new FbParameter("@uniMed", UNI_MED),
                                new FbParameter("@existGral", EXIST_GRAL),
                                new FbParameter("@exist", EXIST),
                                new FbParameter("@fechaElab", FECHA_ELAB),
                                new FbParameter("@cveFolio", CVE_FOLIO),
                                new FbParameter("@signo", SIGNO),
                                new FbParameter("@costoPromIni", COSTO_PROM),
                                new FbParameter("@costoPromFin", COSTO_PROM),
                                new FbParameter("@costoPromGral", COSTO_PROM)
                            );

                            conect.ExecuteNonQuery(sqlActualizarExistenciaAlmacen,
                                new FbParameter("@existNueva", EXIST),
                                new FbParameter("@cveArt", CVE_ART),
                                new FbParameter("@numAlm", NUM_ALM));

                            conect.ExecuteNonQuery(sqlActualizarExistenciaGeneral,
                                new FbParameter("@existGralNueva", EXIST_GRAL),
                                new FbParameter("@cveArt", CVE_ART));

                            conect.ExecuteNonQuery(sqlActualizarUltimaClaveMovimiento,
                                new FbParameter("@numMov", NUM_MOV));

                            conect.ExecuteNonQuery(sqlActualizarUltimoFolio,
                                new FbParameter("@cveFolio", CVE_FOLIO));

                            DESCR_O = conect.ExecuteScalar(sqlLeerDescripcionConcepto,
                                new FbParameter("@cveCptoRev", CVE_CPTO_REV))?.ToString() ?? "NA";

                            conect.ExecuteNonQuery(sqlInsertarBitacoraCancelacion,
                                new FbParameter("@cveBita", Convert.ToInt32(conect.ExecuteScalar($"SELECT COALESCE(MAX(CVE_BITA), 0) + 1 FROM {DVK_BITA}"))),
                                new FbParameter("@cveClie", CVE_CLPV),
                                new FbParameter("@cveCampania", 0),
                                new FbParameter("@cveActividad", CVE_CPTO_REV),
                                new FbParameter("@fechaHora", FECHA_MOV),
                                new FbParameter("@cveUsuario", 1),
                                new FbParameter("@observaciones", $"Cancelación de movimiento {REFER} {CVE_DOC}"),
                                new FbParameter("@status", "C"),
                                new FbParameter("@nomUsuario", Environment.MachineName)
                            );

                            Cell cell16 = new Cell(1, 1)
                                    .SetBorder(Border.NO_BORDER)
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .Add(new Paragraph(NUM_ALM.ToString())
                                        .SetFont(TipoLetra3)
                                        .SetFontSize(10));
                            pdfPtable4.AddCell(cell16);

                            Cell cell17 = new Cell(1, 1)
                                    .SetBorder(Border.NO_BORDER)
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .Add(new Paragraph($"{CVE_CPTO_REV} - {DESCR_O}")
                                        .SetFont(TipoLetra3)
                                        .SetFontSize(10));
                            pdfPtable4.AddCell(cell17);

                            Cell cell18 = new Cell(1, 1)
                                    .SetBorder(Border.NO_BORDER)
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .Add(new Paragraph((CANT * SIGNO).ToString("0.00"))
                                        .SetFont(TipoLetra3)
                                        .SetFontSize(10));
                            pdfPtable4.AddCell(cell18);

                            Cell cell19 = new Cell(1, 1)
                                    .SetBorder(Border.NO_BORDER)
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .Add(new Paragraph(CVE_ART)
                                        .SetFont(TipoLetra3)
                                        .SetFontSize(10));
                            pdfPtable4.AddCell(cell19);

                            Cell cell20 = new Cell(1, 1)
                                    .SetBorder(Border.NO_BORDER)
                                    .SetTextAlignment(TextAlignment.CENTER)
                                    .Add(new Paragraph(DESCR)
                                        .SetFont(TipoLetra3)
                                        .SetFontSize(10));
                            pdfPtable4.AddCell(cell20);
                        }

                        try
                        {
                            document.SetMargins(30f, 30f, 30f, 30f);

                            if (instance != null)
                                document.Add(instance);

                            document.Add(pdfPtable1);
                            document.Add(pdfPtable2);
                            document.Add(pdfPtable3);
                            document.Add(pdfPtable4);

                            Process.Start(rutaPDF);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("No se pudo generar el documento PDF\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("Error durante cancelación:\n" + ex.Message);
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