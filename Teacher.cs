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
    public partial class Teacher : Form
    {
        public Teacher()
        {
            InitializeComponent();
        }

        private void сохранитьToolStripButton4_Click(object sender, EventArgs e)
        {

        }

        private void Teacher_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.View". При необходимости она может быть перемещена или удалена.
            this.viewTableAdapter.Fill(this.schoolFoodDataSet.View);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.View10". При необходимости она может быть перемещена или удалена.
            this.view10TableAdapter.Fill(this.schoolFoodDataSet.View10);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.View8". При необходимости она может быть перемещена или удалена.
            this.view8TableAdapter.Fill(this.schoolFoodDataSet.View8);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Справочник_типов_питания". При необходимости она может быть перемещена или удалена.
            this.справочник_типов_питанияTableAdapter.Fill(this.schoolFoodDataSet.Справочник_типов_питания);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Питание_учеников". При необходимости она может быть перемещена или удалена.
            this.питание_учениковTableAdapter.Fill(this.schoolFoodDataSet.Питание_учеников);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Справочник_классов". При необходимости она может быть перемещена или удалена.
            this.справочник_классовTableAdapter.Fill(this.schoolFoodDataSet.Справочник_классов);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Ученики_школы". При необходимости она может быть перемещена или удалена.
            this.ученики_школыTableAdapter.Fill(this.schoolFoodDataSet.Ученики_школы);

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            dataGridView1.Focus();
            dataGridView1.Refresh();
            dataGridView1.Update();
            dataGridView1.EndEdit();
            bindingNavigator1.BindingSource.EndEdit();
            this.ученики_школыTableAdapter.Update(this.schoolFoodDataSet.Ученики_школы);
            const string message = "Данные сохранены!";
            const string caption = "Учащиеся";
            var result = MessageBox.Show(message, caption,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
           
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            dataGridView1.Focus();
            dataGridView1.Refresh();
            dataGridView1.Update();
            dataGridView1.EndEdit();
            bindingNavigator1.BindingSource.EndEdit();
            this.питание_учениковTableAdapter.Update(this.schoolFoodDataSet.Питание_учеников);
            const string message = "Данные сохранены!";
            const string caption = "Питание учащиехся";
            var result = MessageBox.Show(message, caption,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
        }

        private void tabPage24_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            this.view10TableAdapter.Fill(this.schoolFoodDataSet.View10);
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            this.viewTableAdapter.Fill(this.schoolFoodDataSet.View);
        }
    }
}
