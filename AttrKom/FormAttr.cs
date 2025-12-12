using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Kompas6API5;
using Kompas6Constants;
using KAPITypes;
using reference = System.Int32;

namespace AttrKom
{
    public partial class FormAttr : Form
    {
        DataGridView DataGridview;
        private KompasObject kompas;
        private ksDocument3D doc3D;
        private ksPart part;
        private ksAttributeObject attr;
        private string attrLibraryFile = string.Empty;

        public FormAttr()
        {
            InitializeComponent();
            attrLibraryFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "document_attr.lat");

            // Изначально кнопка недоступна
            toolStripButton1.Enabled = false;
        }

        private void Settings_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // Передаем текущее значение из toolStripTextBox1 в форму операций
            string currentValue = toolStripTextBox1.Text ?? "";
            OperationForm operationForm = new OperationForm(currentValue);

            if (operationForm.ShowDialog() == DialogResult.OK)
            {
                // Если форма закрыта с результатом OK, обновляем toolStripTextBox1
                toolStripTextBox1.Text = operationForm.SelectedCodes;
            }
        }

        private void Write_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что есть имя атрибута
                if (string.IsNullOrEmpty(textBoxAttr.Text))
                {
                    MessageBox.Show("Не выбран атрибут для записи.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Подключаемся к KOMPAS
                if (!ConnectToKompas())
                {
                    MessageBox.Show("KOMPAS-3D не запущен или документ не открыт.\nЗапустите KOMPAS-3D и откройте 3D документ (деталь или сборку).",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Получаем ссылку на 3D документ
                reference docRef = doc3D.reference;
                string attrName = textBoxAttr.Text;
                string newValue = toolStripTextBox1.Text ?? "";

                // Используем метод записи атрибута
                bool success = SetDocumentAttribute(docRef, attrName, newValue);

                if (success)
                {
                    // Обновляем отображение в DataGridView
                    if (DataGridview != null)
                    {
                        foreach (DataGridViewRow row in DataGridview.Rows)
                        {
                            if (row.Cells["Attr"].Value?.ToString() == attrName)
                            {
                                row.Cells["Value"].Value = newValue;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"Не удалось записать значение атрибута '{attrName}'.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при записи атрибута:\n{0}", ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Apply_Click(object sender, EventArgs e)
        {
            try
            {
                // Подключаемся к KOMPAS
                if (!ConnectToKompas())
                {
                    MessageBox.Show("KOMPAS-3D не запущен или документ не открыт.\nЗапустите KOMPAS-3D и откройте 3D документ (деталь или сборку).",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var rows = Rows.Load("Rows.xml");

                // Если DGV не создан — создаём и настраиваем
                if (DataGridview == null)
                {
                    DataGridview = new DataGridView();
                    DataGridview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    DataGridview.AllowUserToAddRows = false;
                    DataGridview.Dock = DockStyle.Fill;
                    DataGridview.RowHeadersVisible = false;
                    DataGridview.AllowUserToResizeRows = false;

                    // Важно: выбор всей строки при клике
                    DataGridview.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    DataGridview.MultiSelect = false;

                    // Подписываемся на события (убираем ранее подписанные на всякий случай)
                    DataGridview.SelectionChanged -= DataGridview_SelectionChanged;
                    DataGridview.SelectionChanged += DataGridview_SelectionChanged;

                    DataGridview.CellClick -= DataGridview_CellClick;
                    DataGridview.CellClick += DataGridview_CellClick;
                }

                // Очищаем таблицу перед заполнением
                DataGridview.Columns.Clear();
                DataGridview.Rows.Clear();

                // Создаём две колонки
                DataGridview.Columns.Add("Attr", "Аттрибут");
                DataGridview.Columns.Add("Value", "Значение");

                // Получаем ссылку на 3D документ
                reference docRef = doc3D.reference;

                // Получаем все существующие атрибуты документа
                Dictionary<string, string> existingAttributes = GetDocumentAttributes(docRef);

                // Сначала создаём недостающие атрибуты
                bool attributesCreated = false;
                foreach (var row in rows.Items)
                {
                    // Проверяем, есть ли атрибут в модели
                    if (!existingAttributes.ContainsKey(row.AttrName))
                    {
                        // Атрибута нет в модели - создаём пустой
                        CreateEmptyAttribute(docRef, row.AttrName);
                        attributesCreated = true;
                    }
                }

                // Если были созданы новые атрибуты, перечитываем список
                if (attributesCreated)
                {
                    existingAttributes = GetDocumentAttributes(docRef);
                }

                // Заполняем DataGridView строками из XML с их значениями
                foreach (var row in rows.Items)
                {
                    string attrValue = "";

                    // Получаем значение атрибута из словаря
                    if (existingAttributes.ContainsKey(row.AttrName))
                    {
                        attrValue = existingAttributes[row.AttrName];
                    }

                    DataGridview.Rows.Add(row.AttrName, attrValue);
                }

                // Удаляем DataGridView из всех контейнеров (если был добавлен раньше)
                if (DataGridview.Parent != null)
                    DataGridview.Parent.Controls.Remove(DataGridview);

                // Очищаем 6 строк первого столбца
                for (int i = 0; i < 6; i++)
                {
                    var control = tableLayoutPanel1.GetControlFromPosition(0, i);
                    if (control != null)
                        tableLayoutPanel1.Controls.Remove(control);
                }

                // Добавляем DGV и растягиваем на 6 строк
                tableLayoutPanel1.Controls.Add(DataGridview, 0, 0);
                tableLayoutPanel1.SetRowSpan(DataGridview, 6);

                DataGridview.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при обработке атрибутов:\n{0}", ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Срабатывает при смене выделения (работает для FullRowSelect)
        private void DataGridview_SelectionChanged(object sender, EventArgs e)
        {
            UpdateTextBoxesFromCurrentRow();
        }

        // Срабатывает при клике на ячейку (на случай, если SelectionChanged не сработало)
        private void DataGridview_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // e.RowIndex может быть -1 (заголовок) — проверим
            if (e.RowIndex < 0) return;
            UpdateTextBoxesFromCurrentRow();
        }

        // Универсальная функция обновления текстбоксов по текущей строке
        private void UpdateTextBoxesFromCurrentRow()
        {
            if (DataGridview == null)
            {
                toolStripButton1.Enabled = false;
                return;
            }

            DataGridViewRow currentRow = null;

            // 1) попробуем CurrentRow
            if (DataGridview.CurrentRow != null)
                currentRow = DataGridview.CurrentRow;
            else if (DataGridview.SelectedRows.Count > 0)
                currentRow = DataGridview.SelectedRows[0];
            else if (DataGridview.SelectedCells.Count > 0)
                currentRow = DataGridview.SelectedCells[0].OwningRow;

            if (currentRow == null)
            {
                textBoxAttr.Text = "";
                toolStripTextBox1.Text = "";
                toolStripButton1.Enabled = false;
                return;
            }

            // Берём значения по именам колонок (безопасно)
            var attrCell = currentRow.Cells["Attr"];
            var valCell = currentRow.Cells["Value"];

            string attrValue = attrCell?.Value?.ToString() ?? "";
            textBoxAttr.Text = attrValue;
            toolStripTextBox1.Text = valCell?.Value?.ToString() ?? "";

            // Включаем кнопку только если выбран атрибут "Технологический маршрут"
            toolStripButton1.Enabled = (attrValue == "Технологический маршрут");
        }

        // Подключение к KOMPAS
        private bool ConnectToKompas()
        {
            try
            {
                if (kompas == null)
                {
                    Type kompasType = Type.GetTypeFromProgID("KOMPAS.Application.5");
                    if (kompasType == null)
                        return false;

                    kompas = (KompasObject)System.Runtime.InteropServices.Marshal.GetActiveObject("KOMPAS.Application.5");
                }

                if (kompas != null)
                {
                    doc3D = (ksDocument3D)kompas.ActiveDocument3D();
                    if (doc3D != null && doc3D.reference != 0)
                    {
                        part = (ksPart)doc3D.GetPart(-1);
                        attr = (ksAttributeObject)kompas.GetAttributeObject();
                        return (part != null && attr != null);
                    }
                }
            }
            catch
            {
                return false;
            }
            return false;
        }

        // Запись значения атрибута документа
        private bool SetDocumentAttribute(reference docRef, string attrName, string attrValue)
        {
            try
            {
                // Создаем итератор атрибутов для 3D документа
                ksIterator iter = (ksIterator)kompas.GetIterator();
                if (iter != null && iter.ksCreateAttrIterator(docRef, 0, 0, 0, 0, 0))
                {
                    // Получаем первый атрибут
                    reference ownerRef = docRef;
                    reference pAttr = iter.ksMoveAttrIterator("F", ref ownerRef);

                    while (pAttr != 0)
                    {
                        try
                        {
                            // Получаем информацию о ключах атрибута и его тип
                            int k1 = 0, k2 = 0, k3 = 0, k4 = 0;
                            double attrTypeNum = 0;
                            attr.ksGetAttrKeysInfo(pAttr, out k1, out k2, out k3, out k4, out attrTypeNum);

                            // Получаем тип атрибута для извлечения имени
                            ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
                            if (type != null)
                            {
                                type.Init();
                                // Получаем информацию о типе атрибута
                                int getTypeResult = attr.ksGetAttrType(attrTypeNum, attrLibraryFile, type);

                                if (getTypeResult == 1)
                                {
                                    string currentAttrName = type.header;

                                    if (currentAttrName == attrName)
                                    {
                                        // Нашли нужный атрибут - записываем новое значение

                                        // Создаем параметры для записи (аналогично DocumentAttributeForm.cs:203-215)
                                        ksUserParam usPar = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
                                        ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
                                        ksDynamicArray valueArr = (ksDynamicArray)kompas.GetDynamicArray(23);

                                        if (usPar != null && item != null && valueArr != null)
                                        {
                                            usPar.Init();
                                            // Связываем массив с параметром (DocumentAttributeForm.cs:208)
                                            usPar.SetUserArray(valueArr);

                                            // Добавляем новое значение в массив (DocumentAttributeForm.cs:209-211)
                                            item.Init();
                                            item.strVal = attrValue;
                                            valueArr.ksAddArrayItem(-1, item);

                                            // Записываем значение атрибута (строка 0, столбец 0)
                                            int setRowResult = attr.ksSetAttrRow(pAttr, 0, 0, 0, usPar, string.Empty);

                                            // Очищаем массив значений
                                            valueArr.ksDeleteArray();

                                            if (setRowResult == 1)
                                            {
                                                return true;
                                            }
                                            else
                                            {
                                                return false;
                                            }
                                        }

                                        break;
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // Игнорируем ошибки при обработке отдельного атрибута
                        }

                        // Переходим к следующему атрибуту
                        ownerRef = docRef;
                        pAttr = iter.ksMoveAttrIterator("N", ref ownerRef);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при записи атрибута:\n{0}", ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // Получение всех атрибутов документа
        private Dictionary<string, string> GetDocumentAttributes(reference docRef)
        {
            Dictionary<string, string> attributes = new Dictionary<string, string>();

            try
            {
                // Создаем параметры и массив ДО цикла (аналогично DocumentAttributeForm.cs:276-288)
                ksUserParam usPar = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
                ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
                ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(23);

                if (usPar == null || item == null || arr == null)
                {
                    MessageBox.Show("Ошибка инициализации параметров.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return attributes;
                }

                usPar.Init();
                usPar.SetUserArray(arr);

                // Заполняем массив начальным пустым значением
                item.Init();
                item.strVal = string.Empty;
                arr.ksAddArrayItem(-1, item);

                // Создаем итератор атрибутов для 3D документа
                ksIterator iter = (ksIterator)kompas.GetIterator();
                if (iter != null && iter.ksCreateAttrIterator(docRef, 0, 0, 0, 0, 0))
                {
                    // Получаем первый атрибут
                    reference ownerRef = docRef;
                    reference pAttr = iter.ksMoveAttrIterator("F", ref ownerRef);

                    while (pAttr != 0)
                    {
                        try
                        {
                            // Получаем информацию о ключах атрибута и его тип
                            int k1 = 0, k2 = 0, k3 = 0, k4 = 0;
                            double attrTypeNum = 0;
                            attr.ksGetAttrKeysInfo(pAttr, out k1, out k2, out k3, out k4, out attrTypeNum);

                            // Получаем тип атрибута для извлечения имени
                            ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
                            if (type != null)
                            {
                                type.Init();
                                // Получаем информацию о типе атрибута
                                int getTypeResult = attr.ksGetAttrType(attrTypeNum, attrLibraryFile, type);

                                if (getTypeResult == 1)
                                {
                                    string attrName = type.header;

                                    if (!string.IsNullOrEmpty(attrName))
                                    {
                                        // Читаем строку атрибута
                                        string attrValue = "";
                                        attr.ksGetAttrRow(pAttr, 0, 0, 0, usPar);

                                        // ВАЖНО: Получаем массив ПОСЛЕ чтения
                                        ksDynamicArray readArr = (ksDynamicArray)usPar.GetUserArray();

                                        if (readArr != null && readArr.ksGetArrayCount() > 0)
                                        {
                                            // Получаем значение из первой колонки
                                            item.Init();
                                            int getItemResult = readArr.ksGetArrayItem(0, item);

                                            if (getItemResult == 1)
                                            {
                                                attrValue = item.strVal ?? "";
                                            }
                                        }

                                        // Добавляем в словарь (даже если значение пустое)
                                        if (!attributes.ContainsKey(attrName))
                                        {
                                            attributes.Add(attrName, attrValue);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // Игнорируем ошибки при чтении отдельного атрибута
                        }

                        // Переходим к следующему атрибуту
                        ownerRef = docRef;
                        pAttr = iter.ksMoveAttrIterator("N", ref ownerRef);
                    }
                }

                // Очищаем массив
                if (arr != null)
                    arr.ksDeleteArray();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при чтении атрибутов:\n{0}", ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return attributes;
        }

        // Создание пустого атрибута
        private void CreateEmptyAttribute(reference docRef, string attrName)
        {
            try
            {
                // Фиксированные ключи для всех атрибутов
                int key1 = 99;
                int key2 = 99;
                int key3 = 99;
                int key4 = 1200;

                // Создаем тип атрибута
                ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
                ksColumnInfoParam col = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);

                if (type != null && col != null)
                {
                    type.Init();
                    col.Init();

                    type.header = attrName;
                    type.rowsCount = 1;
                    type.flagVisible = true;
                    type.password = string.Empty;
                    type.key1 = key1;
                    type.key2 = key2;
                    type.key3 = key3;
                    type.key4 = key4;

                    ksDynamicArray arr = (ksDynamicArray)type.GetColumns();
                    if (arr != null)
                    {
                        col.header = "Значение";
                        col.type = ldefin2d.STRING_ATTR_TYPE;
                        col.key = 0;
                        col.def = "";
                        col.flagEnum = false;
                        arr.ksAddArrayItem(-1, col);

                        // Создаем тип атрибута
                        double numbType = attr.ksCreateAttrType(type, attrLibraryFile);

                        if (numbType > 0)
                        {
                            // Создаем параметры атрибута
                            ksAttributeParam attrPar = (ksAttributeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_Attribute);
                            ksUserParam usPar = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
                            ksDynamicArray fVisibl = (ksDynamicArray)kompas.GetDynamicArray(23);
                            ksDynamicArray colKeys = (ksDynamicArray)kompas.GetDynamicArray(23);

                            if (attrPar != null && usPar != null && fVisibl != null && colKeys != null)
                            {
                                attrPar.Init();
                                usPar.Init();
                                attrPar.SetValues(usPar);
                                attrPar.SetColumnKeys(colKeys);
                                attrPar.SetFlagVisible(fVisibl);
                                attrPar.key1 = key1;
                                attrPar.key2 = key2;
                                attrPar.key3 = key3;
                                attrPar.key4 = key4;
                                attrPar.password = string.Empty;

                                ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
                                ksDynamicArray valueArr = (ksDynamicArray)kompas.GetDynamicArray(23);

                                if (item != null && valueArr != null)
                                {
                                    usPar.SetUserArray(valueArr);
                                    item.Init();
                                    item.strVal = "";
                                    valueArr.ksAddArrayItem(-1, item);

                                    item.Init();
                                    item.uCharVal = 1;
                                    fVisibl.ksAddArrayItem(-1, item);

                                    // Создаем атрибут документа
                                    attr.ksCreateAttr(docRef, attrPar, numbType, attrLibraryFile);

                                    valueArr.ksDeleteArray();
                                }

                                fVisibl.ksDeleteArray();
                                colKeys.ksDeleteArray();
                            }
                        }

                        arr.ksDeleteArray();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при создании атрибута '{0}':\n{1}", attrName, ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
