using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AttrKom
{
    public partial class OperationForm : Form
    {
        public string SelectedCodes { get; private set; }
        private string initialValue;

        public OperationForm(string currentValue = "")
        {
            InitializeComponent();
            initialValue = currentValue;
            InitializeDataGridView();
            AdjustFormHeight();
            SetCheckboxesFromValue(currentValue);
        }

        private void InitializeDataGridView()
        {
            try
            {
                // Путь к XML файлу
                string xmlFilePath = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                    "Operations.xml");

                // Загружаем данные из XML
                Operations operations = Operations.Load(xmlFilePath);

                // Заполняем DataGridView данными из XML
                foreach (var operation in operations.Items)
                {
                    dataGridView1.Rows.Add(false, operation.Name, operation.Code);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных операций:\n{ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AdjustFormHeight()
        {
            // Вычисляем необходимую высоту для формы
            int headerHeight = dataGridView1.ColumnHeadersHeight;
            int rowHeight = dataGridView1.RowTemplate.Height;
            int rowCount = dataGridView1.Rows.Count;

            // Высота всех строк + заголовок
            int gridHeight = headerHeight + (rowHeight * rowCount);

            // Высота панели с кнопками
            int buttonPanelHeight = panelButtons.Height;

            // Добавляем дополнительные отступы (для рамки формы, заголовка окна и т.д.)
            int formBorderHeight = this.Height - this.ClientSize.Height;

            // Устанавливаем высоту формы (grid + панель кнопок + рамка + отступ)
            this.Height = gridHeight + buttonPanelHeight + formBorderHeight + 10;

            // Отключаем возможность изменения размера формы
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
        }

        private void SetCheckboxesFromValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return;

            // Разбиваем строку на отдельные коды (разделитель - запятая с пробелом или без)
            string[] codes = value.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);

            // Проходим по всем строкам DataGridView
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string rowCode = row.Cells["ColumnCode"].Value?.ToString();

                if (!string.IsNullOrEmpty(rowCode))
                {
                    // Проверяем, есть ли код этой строки в списке кодов
                    bool shouldBeChecked = codes.Any(c => c.Trim() == rowCode.Trim());

                    // Устанавливаем значение checkbox
                    row.Cells["ColumnStatus"].Value = shouldBeChecked;
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            // Собираем коды из строк с отмеченным checkbox
            List<string> selectedCodes = new List<string>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Проверяем значение checkbox в первой колонке (ColumnStatus)
                bool isChecked = row.Cells["ColumnStatus"].Value != null &&
                                 (bool)row.Cells["ColumnStatus"].Value;

                if (isChecked)
                {
                    // Получаем код из третьей колонки (ColumnCode)
                    string code = row.Cells["ColumnCode"].Value?.ToString();
                    if (!string.IsNullOrEmpty(code))
                    {
                        selectedCodes.Add(code);
                    }
                }
            }

            // Объединяем коды через запятую
            SelectedCodes = string.Join(", ", selectedCodes);
        }
    }
}
