using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Школьное_питание
{
    public partial class Cooker : Form
    {
        public Cooker()
        {
            InitializeComponent();
        }

        private void Close_Pr_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void Cooker_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet3.View10". При необходимости она может быть перемещена или удалена.
            this.view10TableAdapter.Fill(this.schoolFoodDataSet3.View10);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet6.View8". При необходимости она может быть перемещена или удалена.
            this.view8TableAdapter.Fill(this.schoolFoodDataSet6.View8);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.View". При необходимости она может быть перемещена или удалена.
            this.viewTableAdapter.Fill(this.schoolFoodDataSet.View);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet5.View61". При необходимости она может быть перемещена или удалена.
            this.view61TableAdapter.Fill(this.schoolFoodDataSet5.View61);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.View51". При необходимости она может быть перемещена или удалена.
            this.view51TableAdapter.Fill(this.schoolFoodDataSet.View51);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet4.View3". При необходимости она может быть перемещена или удалена.
            this.view3TableAdapter.Fill(this.schoolFoodDataSet4.View3);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.View41". При необходимости она может быть перемещена или удалена.
            this.view41TableAdapter.Fill(this.schoolFoodDataSet.View41);

            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet3.View4". При необходимости она может быть перемещена или удалена.
            this.view4TableAdapter.Fill(this.schoolFoodDataSet3.View4);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet3.View2". При необходимости она может быть перемещена или удалена.
            this.view2TableAdapter.Fill(this.schoolFoodDataSet3.View2);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.К_расходу". При необходимости она может быть перемещена или удалена.
            this.к_расходуTableAdapter.Fill(this.schoolFoodDataSet.К_расходу);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Расход_товаров". При необходимости она может быть перемещена или удалена.
            this.расход_товаровTableAdapter.Fill(this.schoolFoodDataSet.Расход_товаров);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.К_поставке_товаров". При необходимости она может быть перемещена или удалена.
            this.к_поставке_товаровTableAdapter.Fill(this.schoolFoodDataSet.К_поставке_товаров);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Поставка_товаров". При необходимости она может быть перемещена или удалена.
            this.поставка_товаровTableAdapter.Fill(this.schoolFoodDataSet.Поставка_товаров);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Затраты_продуктов". При необходимости она может быть перемещена или удалена.
            this.затраты_продуктовTableAdapter.Fill(this.schoolFoodDataSet.Затраты_продуктов);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Питание_учеников". При необходимости она может быть перемещена или удалена.
            this.питание_учениковTableAdapter.Fill(this.schoolFoodDataSet.Питание_учеников);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Питание_учителей". При необходимости она может быть перемещена или удалена.
            this.питание_учителейTableAdapter.Fill(this.schoolFoodDataSet.Питание_учителей);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Ученики_школы". При необходимости она может быть перемещена или удалена.
            this.ученики_школыTableAdapter.Fill(this.schoolFoodDataSet.Ученики_школы);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Учителя_школы". При необходимости она может быть перемещена или удалена.
            this.учителя_школыTableAdapter.Fill(this.schoolFoodDataSet.Учителя_школы);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Справочник_классов". При необходимости она может быть перемещена или удалена.
            this.справочник_классовTableAdapter.Fill(this.schoolFoodDataSet.Справочник_классов);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Справочник_продуктов". При необходимости она может быть перемещена или удалена.
            this.справочник_продуктовTableAdapter.Fill(this.schoolFoodDataSet.Справочник_продуктов);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Справочник_типов_питания". При необходимости она может быть перемещена или удалена.
            this.справочник_типов_питанияTableAdapter.Fill(this.schoolFoodDataSet.Справочник_типов_питания);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "schoolFoodDataSet.Справочник_поставщиков". При необходимости она может быть перемещена или удалена.
            this.справочник_поставщиковTableAdapter.Fill(this.schoolFoodDataSet.Справочник_поставщиков);

        }

        private void сохранитьToolStripButton9_Click(object sender, EventArgs e)
        {
            dataGridView1.Focus();
            dataGridView1.Refresh();
            dataGridView1.Update();
            dataGridView1.EndEdit();
            bindingNavigator10.BindingSource.EndEdit();
            this.поставка_товаровTableAdapter.Update(this.schoolFoodDataSet.Поставка_товаров);
            const string message = "Данные накладной сохранены!";
            const string caption = "Поставка продуктов";
            var result = MessageBox.Show(message, caption,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
        }

        private void сохранитьToolStripButton10_Click(object sender, EventArgs e)
        {
            dataGridView11.Focus();
            dataGridView11.Refresh();
            dataGridView11.Update();
            dataGridView11.EndEdit();
            bindingNavigator11.BindingSource.EndEdit();
            this.к_поставке_товаровTableAdapter.Update(this.schoolFoodDataSet.К_поставке_товаров);
            const string message = "Данные продуктах в накладной сохранены!";
            const string caption = "Поставка продуктов";
            var result = MessageBox.Show(message, caption,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            dataGridView12.Focus();
            dataGridView12.Refresh();
            dataGridView12.Update();
            dataGridView12.EndEdit();
            bindingNavigator12.BindingSource.EndEdit();
            this.расход_товаровTableAdapter.Update(this.schoolFoodDataSet.Расход_товаров);
            const string message = "Данные расходной накладной сохранены!";
            const string caption = "Расход продуктов";
            var result = MessageBox.Show(message, caption,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
        }

        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            dataGridView13.Focus();
            dataGridView13.Refresh();
            dataGridView13.Update();
            dataGridView13.EndEdit();
            bindingNavigator13.BindingSource.EndEdit();
            this.к_расходуTableAdapter.Update(this.schoolFoodDataSet.К_расходу);
            const string message = "Данные расходе продуктов сохранены!";
            const string caption = "Расход продуктов";
            var result = MessageBox.Show(message, caption,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
        }

        private void сохранитьToolStripButton8_Click(object sender, EventArgs e)
        {
            dataGridView9.Focus();
            dataGridView9.Refresh();
            dataGridView9.Update();
            dataGridView9.EndEdit();
            bindingNavigator10.BindingSource.EndEdit();
            this.затраты_продуктовTableAdapter.Update(this.schoolFoodDataSet.Затраты_продуктов);
            const string message = "Данные сохранены!";
            const string caption = "Затраты продуктов на порцию";
            var result = MessageBox.Show(message, caption,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
        }

        private void сохранитьToolStripButton6_Click(object sender, EventArgs e)
        {
            dataGridView7.Focus();
            dataGridView7.Refresh();
            dataGridView7.Update();
            dataGridView7.EndEdit();
            bindingNavigator7.BindingSource.EndEdit();
            this.питание_учителейTableAdapter.Update(this.schoolFoodDataSet.Питание_учителей);
            const string message = "Данные сохранены!";
            const string caption = "Питание учителей школы";
            var result = MessageBox.Show(message, caption,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
        }

        private void сохранитьToolStripButton7_Click(object sender, EventArgs e)
        {
            dataGridView8.Focus();
            dataGridView8.Refresh();
            dataGridView8.Update();
            dataGridView8.EndEdit();
            bindingNavigator8.BindingSource.EndEdit();
            this.питание_учениковTableAdapter.Update(this.schoolFoodDataSet.Питание_учеников);
            const string message = "Данные сохранены!";
            const string caption = "Питание учащихся школы";
            var result = MessageBox.Show(message, caption,
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Information);
        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            this.view51TableAdapter.Fill(this.schoolFoodDataSet.View51);
        }

        private void toolStripButton16_Click(object sender, EventArgs e)
        {
            this.view61TableAdapter.Fill(this.schoolFoodDataSet5.View61);
        }

        private void toolStripButton17_Click(object sender, EventArgs e)
        {
            this.view2TableAdapter.Fill(this.schoolFoodDataSet3.View2);
        }

        private void toolStripButton18_Click(object sender, EventArgs e)
        {
            this.view41TableAdapter.Fill(this.schoolFoodDataSet.View41);
        }

        private void toolStripButton19_Click(object sender, EventArgs e)
        {
            this.view3TableAdapter.Fill(this.schoolFoodDataSet4.View3);
        }

        private void toolStripButton20_Click(object sender, EventArgs e)
        {
            this.viewTableAdapter.Fill(this.schoolFoodDataSet.View);
        }

        private void toolStripButton21_Click(object sender, EventArgs e)
        {
            this.view8TableAdapter.Fill(this.schoolFoodDataSet6.View8);
        }

        private void toolStripButton22_Click(object sender, EventArgs e)
        {
            //Экспорт в Excel  
            Excel.Application exApp = new Excel.Application();

            exApp.Visible = true;
            exApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

            workSheet.Cells[1, 1] = "Ведомость о питании по классам";
            workSheet.Cells[3, 1] = "Месяц";
            workSheet.Cells[3, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 2] = "Год";
            workSheet.Cells[3, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 3] = "Название класса";
            workSheet.Cells[3, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 4] = "Название типа питания";
            workSheet.Cells[3, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 5] = "Количество";
            workSheet.Cells[3, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 6] = "Цена питания";
            workSheet.Cells[3, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            int rowExcel = 4;
            for (int i = 0; i < dataGridView17.Rows.Count-1; i++)
            {
                for (int j = 0; j < dataGridView17.Columns.Count; j++)
                {
                    workSheet.Cells[i + rowExcel, j + 1] = dataGridView17.Rows[i].Cells[j].Value.ToString();
                    workSheet.Cells[i + rowExcel, j + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }
            }
            workSheet.SaveAs("Ведомость о питании по классам.xlsx");
            MessageBox.Show("Экспорт данных завершен...", "Школьное питание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            exApp.Quit();
        }

        private void fillByToolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                this.view51TableAdapter.FillBy(this.schoolFoodDataSet.View51, ((decimal)(System.Convert.ChangeType(param1ToolStripTextBox1.Text, typeof(decimal)))));
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillBy2ToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.view51TableAdapter.FillBy2(this.schoolFoodDataSet.View51, param1ToolStripTextBox3.Text);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.view51TableAdapter.Fill(this.schoolFoodDataSet.View51);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }
        private void fillToolStripButton_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.view61TableAdapter.Fill(this.schoolFoodDataSet.View61);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillFIO1ToolStripButton_Click_1(object sender, EventArgs e)
        {

        }

        private void toolStripButton23_CheckStateChanged(object sender, EventArgs e)
        {

        }

        private void fillToolStripButton_Click_2(object sender, EventArgs e)
        {
            try
            {
                this.view61TableAdapter.Fill(this.schoolFoodDataSet.View61);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillBy3ToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.view51TableAdapter.FillBy3(this.schoolFoodDataSet.View51, ((decimal)(System.Convert.ChangeType(param1ToolStripTextBox1.Text, typeof(decimal)))), ((decimal)(System.Convert.ChangeType(param2ToolStripTextBox.Text, typeof(decimal)))));
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void toolStripButton24_Click(object sender, EventArgs e)
        {
            //Экспорт в Excel  
            Excel.Application exApp = new Excel.Application();

            exApp.Visible = true;
            exApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

            workSheet.Cells[1, 1] = "Ведомость о питании по педагогам";
            workSheet.Cells[3, 1] = "Месяц";
            workSheet.Cells[3, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 2] = "Год";
            workSheet.Cells[3, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 3] = "ФИО учителя";
            workSheet.Cells[3, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 4] = "Название типа питания";
            workSheet.Cells[3, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 5] = "Количество";
            workSheet.Cells[3, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 6] = "Цена питания";
            workSheet.Cells[3, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            int rowExcel = 4;
            for (int i = 0; i < dataGridView18.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView18.Columns.Count; j++)
                {
                    workSheet.Cells[i + rowExcel, j + 1] = dataGridView18.Rows[i].Cells[j].Value.ToString();
                    workSheet.Cells[i + rowExcel, j + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }
            }
            workSheet.SaveAs("Ведомость о питании по преподавателям.xlsx");
            MessageBox.Show("Экспорт данных завершен...", "Школьное питание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            exApp.Quit();
        }

        private void toolStripButton25_Click(object sender, EventArgs e)
        {
            //Экспорт в Excel  
            Excel.Application exApp = new Excel.Application();

            exApp.Visible = true;
            exApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

            workSheet.Cells[1, 1] = "Отчет по затратам на каждый тип питания";
            workSheet.Cells[3, 1] = "Дата затрат";
            workSheet.Cells[3, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 2] = "Наименование продукта";
            workSheet.Cells[3, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 3] = "Количество на порцию";
            workSheet.Cells[3, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 4] = "Единица измерения";
            workSheet.Cells[3, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 5] = "Цена";
            workSheet.Cells[3, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 6] = "Название типа питания";
            workSheet.Cells[3, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            

            int rowExcel = 4;
            for (int i = 0; i < dataGridView15.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView15.Columns.Count; j++)
                {
                    workSheet.Cells[i + rowExcel, j + 1] = dataGridView15.Rows[i].Cells[j].Value.ToString();
                    workSheet.Cells[i + rowExcel, j + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }
            }
            workSheet.SaveAs("Отчет по затратам на каждый тип питания.xlsx");
            MessageBox.Show("Экспорт данных завершен...", "Школьное питание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            exApp.Quit();
        }

        private void toolStripButton26_Click(object sender, EventArgs e)
        {
            //Экспорт в Excel  
            Excel.Application exApp = new Excel.Application();

            exApp.Visible = true;
            exApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

            workSheet.Cells[1, 1] = "Отчет о приходе продуктов";
            workSheet.Cells[3, 1] = "Дата поставки";
            workSheet.Cells[3, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 2] = "Наименование постащика";
            workSheet.Cells[3, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 3] = "Название продукта";
            workSheet.Cells[3, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 4] = "Единица измерения";
            workSheet.Cells[3, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 5] = "Цена за единицу измерения";
            workSheet.Cells[3, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 6] = "Количество";
            workSheet.Cells[3, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 7] = "Итого";
            workSheet.Cells[3, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            int rowExcel = 4;
            for (int i = 0; i < dataGridView14.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView14.Columns.Count; j++)
                {
                    workSheet.Cells[i + rowExcel, j + 1] = dataGridView14.Rows[i].Cells[j].Value.ToString();
                    workSheet.Cells[i + rowExcel, j + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }
            }
            workSheet.SaveAs("Отчет о приходе продуктов.xlsx");
            MessageBox.Show("Экспорт данных завершен...", "Школьное питание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            exApp.Quit();
        }

        private void toolStripButton27_Click(object sender, EventArgs e)
        {
            //Экспорт в Excel  
            Excel.Application exApp = new Excel.Application();

            exApp.Visible = true;
            exApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

            workSheet.Cells[1, 1] = "Отчет о расходе продуктов";
            workSheet.Cells[3, 1] = "Дата расхода";
            workSheet.Cells[3, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 2] = "Наименование продукта";
            workSheet.Cells[3, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 3] = "Цена за единицу измерения";
            workSheet.Cells[3, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 4] = "Единица измерения";
            workSheet.Cells[3, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 5] = "Количество";
            workSheet.Cells[3, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 6] = "Общая сумма";
            workSheet.Cells[3, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            
            int rowExcel = 4;
            for (int i = 0; i < dataGridView16.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView16.Columns.Count; j++)
                {
                    workSheet.Cells[i + rowExcel, j + 1] = dataGridView16.Rows[i].Cells[j].Value.ToString();
                    workSheet.Cells[i + rowExcel, j + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }
            }
            workSheet.SaveAs("Отчет о расходе продуктов.xlsx");
            MessageBox.Show("Экспорт данных завершен...", "Школьное питание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            exApp.Quit();
        }

        private void toolStripButton28_Click(object sender, EventArgs e)
        {
            //Экспорт в Excel  
            Excel.Application exApp = new Excel.Application();

            exApp.Visible = true;
            exApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

            workSheet.Cells[1, 1] = "Табель посещений столовой учащимися";
            workSheet.Cells[3, 1] = "Дата";
            workSheet.Cells[3, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 2] = "Название класса";
            workSheet.Cells[3, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 3] = "ФИО ученика";
            workSheet.Cells[3, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 4] = "Название типа питания";
            workSheet.Cells[3, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            int rowExcel = 4;
            for (int i = 0; i < dataGridView19.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView19.Columns.Count; j++)
                {
                    workSheet.Cells[i + rowExcel, j + 1] = dataGridView19.Rows[i].Cells[j].Value.ToString();
                    workSheet.Cells[i + rowExcel, j + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }
            }
            workSheet.SaveAs("Табель посещений столовой учащимися.xlsx");
            MessageBox.Show("Экспорт данных завершен...", "Школьное питание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            exApp.Quit();
        }

        private void toolStripButton29_Click(object sender, EventArgs e)
        {
            //Экспорт в Excel  
            Excel.Application exApp = new Excel.Application();

            exApp.Visible = true;
            exApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

            workSheet.Cells[1, 1] = "Табель посещений столовой преподавателями";
            workSheet.Cells[3, 1] = "Дата";
            workSheet.Cells[3, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 2] = "ФИО учителя";
            workSheet.Cells[3, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            workSheet.Cells[3, 3] = "Название типа питания";
            workSheet.Cells[3, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            int rowExcel = 4;
            for (int i = 0; i < dataGridView20.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView20.Columns.Count; j++)
                {
                    workSheet.Cells[i + rowExcel, j + 1] = dataGridView20.Rows[i].Cells[j].Value.ToString();
                    workSheet.Cells[i + rowExcel, j + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }
            }
            workSheet.SaveAs("Табель посещений столовой преподавателями.xlsx");
            MessageBox.Show("Экспорт данных завершен...", "Школьное питание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            exApp.Quit();
        }
    }
}
