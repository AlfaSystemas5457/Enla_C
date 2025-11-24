using Enla_C.DB;
using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
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

namespace Enla_C
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void generate_policies_Click(object sender, RoutedEventArgs e)
        {
            ToggleView(generatePolicesControl);
        }

        private void cancel_movs_Click(object sender, RoutedEventArgs e)
        {
            ToggleView(cancelMovs);
        }

        private void config_Click(object sender, RoutedEventArgs e)
        {
            ToggleView(configViewControl);
        }

        private void ToggleView(UserControl userControl)
        {
            if (userControl == null)
                return;

            UserControl[] userControlList = new UserControl[]
            {
                aboutViewControl,
                configViewControl,
                generatePolicesControl,
                cancelMovs
            };

            bool anyVisible = false;

            foreach (UserControl uc in userControlList)
            {
                if (uc == null)
                    continue;

                if (uc == userControl)
                {
                    if (uc.Visibility == Visibility.Visible)
                    {
                        uc.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        uc.Visibility = Visibility.Visible;
                        anyVisible = true;
                    }
                }
                else
                {
                    uc.Visibility = Visibility.Collapsed;
                }
            }

            if (!anyVisible)
            {
                aboutViewControl.Visibility = Visibility.Visible;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            FbConnection.ClearAllPools();
        }
    }
}
