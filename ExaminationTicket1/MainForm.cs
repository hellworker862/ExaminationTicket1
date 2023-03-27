using ExaminationTicket1.Modals;
using ExaminationTicket1.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ExaminationTicket1
{
    public partial class MainForm : Form
    {
        private decimal _costCurWindowType;
        private decimal _sum;

        public MainForm()
        {
            InitializeComponent();

            var windowServices = new List<WindowService>();
            windowServices.Add(new WindowService("окна", 192.5m));
            windowServices.Add(new WindowService("балконы", 213.9m));
            windowServices.Add(new WindowService("двери", 342.5m));

            comboBoxWindowServices.ValueMember = "CostSqM";
            comboBoxWindowServices.DisplayMember = "Title";
            comboBoxWindowServices.DataSource = windowServices;
            radioButtonBlind.Checked = true;
        }

        private void buttonAddPhoto_Click(object sender, EventArgs e)
        {
            if (pictureBoxCompanyLogo.Image == null)
            {
                pictureBoxCompanyLogo.Image = Resources.Window1;
                buttonAddPhoto.Text = "Убрать фото";
            }
            else
            {
                pictureBoxCompanyLogo.Image = null;
                buttonAddPhoto.Text = "Добавить фото";
            }

        }

        private void buttonCalculate_Click(object sender, EventArgs e)
        {
            var height = numericUpDownHeight.Value;
            var width = numericUpDownWidth.Value;
            var costWindowService = (decimal)comboBoxWindowServices.SelectedValue;
            errorProvider.Clear();


            if (height <= 0)
                errorProvider.SetError(numericUpDownHeight, "Высота должна быть больше 0");
            if (width <= 0)
                errorProvider.SetError(numericUpDownWidth, "Ширина должна быть больше 0");
            if (errorProvider.GetError(numericUpDownHeight) != string.Empty ||
                errorProvider.GetError(numericUpDownWidth) != string.Empty)
                return;

            try
            {
                var sum = Culculate(height, width, _costCurWindowType, costWindowService);
                listBoxResult.Items.Clear();
                listBoxResult.Items.Add($"Итоговая сумма: {sum}");
                _sum = sum;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public decimal Culculate(decimal height, decimal width, decimal costWindowType, decimal costWindowService)
        {
            if (height <= 0)
                throw new ArgumentException($"Аргумент {nameof(height)} не может быть меньше или равно 0");
            if (width <= 0)
                throw new ArgumentException($"Аргумент {nameof(width)} не может быть меньше или равно 0");
            if (costWindowType <= 0)
                throw new ArgumentException($"Аргумент {nameof(costWindowType)} не может быть меньше или равно 0");
            if (costWindowService <= 0)
                throw new ArgumentException($"Аргумент {nameof(costWindowService)} не может быть меньше или равно 0");

            //приводим к метрам
            var area = (width * height) / 10000;
            var sum = Math.Round((area * costWindowService) + costWindowType, 2);

            return sum;
        }

        private void radioButtonBlind_CheckedChanged(object sender, EventArgs e)
        {
            _costCurWindowType = 1000;
        }

        private void radioButtonPivot_CheckedChanged(object sender, EventArgs e)
        {
            _costCurWindowType = 3400.5m;
        }

        private void radioButtonFlip_CheckedChanged(object sender, EventArgs e)
        {
            _costCurWindowType = 2560;
        }

        private void radioButtonTransom_CheckedChanged(object sender, EventArgs e)
        {
            _costCurWindowType = 7900.9m;
        }

        private void radioButtonSliding_CheckedChanged(object sender, EventArgs e)
        {
            _costCurWindowType = 6210.5m;
        }

        private async void buttonCreateReport_Click(object sender, EventArgs e)
        {
            if (_sum != 0)
            {
                var id = Guid.NewGuid().ToString();
                var date = DateTime.Now.ToString("d");
                var title = ((WindowService)comboBoxWindowServices.SelectedItem).Title;
                var sum = _sum.ToString().Replace(',', '.');
                var data = new Dictionary<string, string>()
                {
                    { "{Уникальный_номер}", id },
                    { "{дата}", date},
                    { "{Товар}", title},
                    { "{итог}", sum }
                };
                try
                {
                    var export = new ExportManager();
                    var fileName = $"{id}_{date}_{_sum}.docx";
                    var savePath = Path.Combine(Environment.CurrentDirectory, "..\\..\\Reporters\\", fileName);
                    await export.ToWordAsync("pattern.docx", data, savePath);
                    MessageBox.Show("Квитанция сохранена");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
