using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using static CheckCheckNET4.MainWindow;

namespace CheckCheckNET4
{
    public class Utils
    {

        //public static string userDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString();

        //public static string dubToolDir = userDocumentsPath + @"\dubtool\";


        public static void listFiles(ListBox listbox, ComboBox combobox)
        {
            listbox.Items.Clear();
            combobox.Items.Clear();

            DirectoryInfo dinfo = new DirectoryInfo(GlobalVariables.dubToolDir);
            if (dinfo.Exists)
            {
                FileInfo[] Files = dinfo.GetFiles("*.xls");
                foreach (FileInfo file in Files)
                {
                    listbox.Items.Add(file.Name);
                    combobox.Items.Add(file.Name);
                }
            }

        }
        // Fyller ut listboks og combobox
        public static void listFilesFromMemoryList(List<string> folderContent, ListBox listbox)
        {
            listbox.Items.Clear();

            string onlyFileName;

            foreach (var file in folderContent)
            {

                onlyFileName = file.Substring(GlobalVariables.dubToolDir.Length, file.Length - GlobalVariables.dubToolDir.Length - 4);
                listbox.Items.Add(onlyFileName);
                //combobox.Items.Add(onlyFileName);
            }
        }



        public static class Prompt
        {
            public static string ShowDialog(string headerText)
            {
                Window window = new Window
                {
                    Width = 400,
                    Height = 150,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen

                };

                Grid grid = new Grid
                {
                    Width = 400,
                    Height = 150

                };


                TextBlock header = new TextBlock { Width = 100, Height = 50 }; ;
                header.Text = headerText;
                TextBox textBox = new TextBox() { Width = 200, Height = 50, FontSize = 30 };

                Button confirmation = new Button() { Content = "Ok", Height = 50, Width = 100 };
                confirmation.Click += (sender, e) => { window.Close(); };
                grid.Children.Add(textBox);
                grid.Children.Add(confirmation);

                window.Content = grid;

                window.ShowDialog();

                return textBox.Text;
            }
        }

        public static void ShowSplashScreen(bool show)
        {
            var window = new Window { Width = 500, Height = 400, Background = Brushes.Blue, WindowStartupLocation = WindowStartupLocation.CenterScreen };

            var grid = new Grid();
            var textBlock = new TextBlock { Text = "Vent litt..", HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, Foreground = Brushes.Wheat, FontSize = 20 };

            window.Content = grid;
            grid.Children.Add(textBlock);

            if (show)
            {
                window.Show();
            }
            else
            {
                window.Hide();
            }



        }

        public static string ToStringAndLower(string s)
        {
            return s.ToString().ToLower();
        }


    }
}
