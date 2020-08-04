using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Школьное_питание
{
    public partial class FormAvtor : Form
    {
        public FormAvtor()
        {
            InitializeComponent();
        }

        private void buttonVxod_Click(object sender, EventArgs e)
        {
            if ((comboBoxUser.SelectedIndex == 0) && (Parol.Text == "111"))
            {
                const string message = "Доступ разрешен!";
                const string caption = "Авторизация";
                var result = MessageBox.Show(message, caption,
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Information);
                Admin f = new Admin();
                f.ShowDialog();               
            }
            else if ((comboBoxUser.SelectedIndex == 1) && (Parol.Text == "222"))
            {
                const string message = "Доступ разрешен!";
                const string caption = "Авторизация";
                var result = MessageBox.Show(message, caption,
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Information);
                Teacher f = new Teacher();
                f.ShowDialog();

            }

            else if ((comboBoxUser.SelectedIndex == 2) && (Parol.Text == "333"))
            {
                const string message = "Доступ разрешен!";
                const string caption = "Авторизация";
                var result = MessageBox.Show(message, caption,
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Information);
                Cooker f = new Cooker();
                f.ShowDialog();
                
            }
            else
            {
                const string message = "Не корректный пароль!";
                const string caption = "Авторизация";
                var result = MessageBox.Show(message, caption,
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Error);
                Parol.Clear();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
           Application.Exit();
           
        }

        private void Parol_KeyDown(object sender, KeyEventArgs e)
        {
            // Проверям нажата ли именно клавиша Enter
            if (e.KeyCode == Keys.Enter)
            {
                buttonVxod.Focus();
            }
        }
    }
}
