﻿using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Akhmetzyanov.xaml
    /// </summary>
    public partial class _4333_Akhmetzyanov : Window
    {
        public _4333_Akhmetzyanov()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Вы нажали на кнопку!");
        }
    }
}
