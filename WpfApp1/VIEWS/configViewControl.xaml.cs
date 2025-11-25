using Enla_C.DB;
using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
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

namespace Enla_C.VIEWS
{
    /// <summary>
    /// Lógica de interacción para UserControl1.xaml
    /// </summary>
    public partial class configViewControl : UserControl
    {
        public generatePolicesControl GeneratePoliciesControl { get; set; }
        ConectDB conect = ConectDB.Instance;

        public configViewControl()
        {
            InitializeComponent();
        }

        private void LoadDataConfig()
        {
            try
            {
                this.conect.LoadSettingsIni();

                textBoxRS.Text = this.conect.RS;
                textBoxServer.Text = this.conect.server;
                textBoxDBAspelSAE.Text = this.conect.baseDatos;
                textBoxNoEmpresa.Text = this.conect.NoEmpresa;
                textBoxUser.Text = this.conect.usuario;
                textBoxPass.Password = this.conect.password;

                textBoxAlm1.Text = this.conect.Alm1;
                textBoxAlm2.Text = this.conect.Alm2;

                checkBoxPermitirMateriaPrima.IsChecked = this.conect.EditaMP;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataConfig();
        }

        private void buttonAccept_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(textBoxRS.Text))
                {
                    MessageBox.Show("El campo razón social es obligatorio.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBoxRS.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBoxServer.Text))
                {
                    MessageBox.Show("El campo nombre de servidor es obligatorio.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBoxServer.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBoxDBAspelSAE.Text))
                {
                    MessageBox.Show("El campo DB ASPEL SAE es obligatorio.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBoxDBAspelSAE.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBoxNoEmpresa.Text))
                {
                    MessageBox.Show("El campo # de empresa SAE es obligatorio.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBoxNoEmpresa.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBoxUser.Text))
                {
                    MessageBox.Show("El campo Usuario es obligatorio.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBoxUser.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBoxPass.Password))
                {
                    MessageBox.Show("El campo contraseña es obligatorio.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBoxPass.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBoxAlm1.Text))
                {
                    MessageBox.Show("El campo almacen 1 es obligatorio.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBoxAlm1.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(textBoxAlm2.Text))
                {
                    MessageBox.Show("El campo almacen 2 es obligatorio.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBoxAlm2.Focus();
                    return;
                }

                FbConnection.ClearAllPools();
                dbParams dbParams = new dbParams()
                {
                    RS = textBoxRS.Text,
                    server = textBoxServer.Text,
                    baseDatos = textBoxDBAspelSAE.Text,
                    NoEmpresa = textBoxNoEmpresa.Text,
                    usuario = textBoxUser.Text,
                    password = textBoxPass.Password,
                    Alm1 = textBoxAlm1.Text,
                    Alm2 = textBoxAlm2.Text,
                    EditaMP = (bool)checkBoxPermitirMateriaPrima.IsChecked
                };

                conect.SaveConfig(dbParams);
                GeneratePoliciesControl?.CambiarAlmacen();
                GeneratePoliciesControl?.loadData();

                using (var conn = this.conect.GetConnectionTest())
                {
                    if (!File.Exists(conect.baseDatos))
                    {
                        GeneratePoliciesControl?.UnloadData();
                        throw new FileNotFoundException("Archivo de base de datos no encontrado.");
                    }

                    conn.Open();
                    MessageBox.Show("Conexión exitosa a la base de datos.", "Éxito", MessageBoxButton.OK, MessageBoxImage.Information);
                    conn.Close();
                }
            }
            catch (FbException ex)
            {
                GeneratePoliciesControl?.UnloadData();
                MessageBox.Show($"No se pudo conectar a la base de datos:\n{ex.Message}", "Error de conexión", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                GeneratePoliciesControl?.UnloadData();
                MessageBox.Show($"Error:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
    }
}
