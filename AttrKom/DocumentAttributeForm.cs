using System;
using System.Windows.Forms;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using KAPITypes;
using reference = System.Int32;

namespace AttrKom
{
    public class DocumentAttributeForm : Form
    {
        private Button btnCreateAttribute;
        private Button btnDeleteAttribute;
        private Button btnShowAttribute;
        private Label lblAttributeValue;
        private KompasObject kompas;
        private ksDocument3D doc3D;
        private ksPart part;
        private ksAttributeObject attr;
        private double createdAttrTypeNumber = 0;
        private string attrLibraryFile = string.Empty;

        public DocumentAttributeForm()
        {
            InitializeComponents();
            attrLibraryFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "document_attr.lat");
        }

        private void InitializeComponents()
        {
            this.Text = "KOMPAS-3D - Атрибуты документа";
            this.Width = 450;
            this.Height = 320;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            btnCreateAttribute = new Button();
            btnCreateAttribute.Text = "Создать атрибут документа 'Наименование'";
            btnCreateAttribute.Width = 350;
            btnCreateAttribute.Height = 50;
            btnCreateAttribute.Top = 20;
            btnCreateAttribute.Left = (this.ClientSize.Width - btnCreateAttribute.Width) / 2;
            btnCreateAttribute.Click += BtnCreateAttribute_Click;
            this.Controls.Add(btnCreateAttribute);

            btnDeleteAttribute = new Button();
            btnDeleteAttribute.Text = "Удалить атрибут документа 'Наименование'";
            btnDeleteAttribute.Width = 350;
            btnDeleteAttribute.Height = 50;
            btnDeleteAttribute.Top = 80;
            btnDeleteAttribute.Left = (this.ClientSize.Width - btnDeleteAttribute.Width) / 2;
            btnDeleteAttribute.Click += BtnDeleteAttribute_Click;
            this.Controls.Add(btnDeleteAttribute);

            btnShowAttribute = new Button();
            btnShowAttribute.Text = "Показать значение атрибута 'Наименование'";
            btnShowAttribute.Width = 350;
            btnShowAttribute.Height = 50;
            btnShowAttribute.Top = 140;
            btnShowAttribute.Left = (this.ClientSize.Width - btnShowAttribute.Width) / 2;
            btnShowAttribute.Click += BtnShowAttribute_Click;
            this.Controls.Add(btnShowAttribute);

            lblAttributeValue = new Label();
            lblAttributeValue.Text = "Значение: не загружено";
            lblAttributeValue.Width = 400;
            lblAttributeValue.Height = 40;
            lblAttributeValue.Top = 210;
            lblAttributeValue.Left = (this.ClientSize.Width - lblAttributeValue.Width) / 2;
            lblAttributeValue.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            lblAttributeValue.BorderStyle = BorderStyle.FixedSingle;
            this.Controls.Add(lblAttributeValue);
        }

        private void BtnCreateAttribute_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ConnectToKompas())
                {
                    MessageBox.Show("KOMPAS-3D не запущен или документ не открыт.\nЗапустите KOMPAS-3D и откройте 3D документ (деталь или сборку).",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                CreateDocumentAttribute();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при создании атрибута:\n{0}", ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDeleteAttribute_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ConnectToKompas())
                {
                    MessageBox.Show("KOMPAS-3D не запущен или документ не открыт.\nЗапустите KOMPAS-3D и откройте 3D документ (деталь или сборку).",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DeleteDocumentAttribute();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при удалении атрибута:\n{0}", ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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

        private void CreateDocumentAttribute()
        {
            // Создаем тип атрибута
            ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)StructType2DEnum.ko_AttributeType);
            ksColumnInfoParam col = (ksColumnInfoParam)kompas.GetParamStruct((short)StructType2DEnum.ko_ColumnInfoParam);

            if (type != null && col != null)
            {
                type.Init();
                col.Init();

                type.header = "Наименование";
                type.rowsCount = 1;
                type.flagVisible = true;
                type.password = string.Empty;
                type.key1 = 100;
                type.key2 = 200;
                type.key3 = 300;
                type.key4 = 0;

                ksDynamicArray arr = (ksDynamicArray)type.GetColumns();
                if (arr != null)
                {
                    col.header = "Значение";
                    col.type = ldefin2d.STRING_ATTR_TYPE;
                    col.key = 0;
                    col.def = "Деталь какая то";
                    col.flagEnum = false;
                    arr.ksAddArrayItem(-1, col);

                    // Создаем тип атрибута
                    double numbType = attr.ksCreateAttrType(type, attrLibraryFile);

                    if (numbType > 0)
                    {
                        createdAttrTypeNumber = numbType;

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
                            attrPar.key1 = 100;
                            attrPar.key2 = 200;
                            attrPar.key3 = 300;
                            attrPar.password = string.Empty;

                            ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
                            ksDynamicArray valueArr = (ksDynamicArray)kompas.GetDynamicArray(23);

                            if (item != null && valueArr != null)
                            {
                                usPar.SetUserArray(valueArr);
                                item.Init();
                                item.strVal = "Деталь какая то";
                                valueArr.ksAddArrayItem(-1, item);

                                item.Init();
                                item.uCharVal = 1;
                                fVisibl.ksAddArrayItem(-1, item);

                                // ВАЖНО: Передаем doc3D.reference для создания атрибута 3D документа
                                // Это создаст атрибут для всей 3D модели
                                reference docRef = doc3D.reference;
                                reference attrRef = attr.ksCreateAttr(docRef, attrPar, numbType, attrLibraryFile);

                                if (attrRef != 0)
                                {
                                    MessageBox.Show("Атрибут 'Наименование' со значением 'Деталь какая то' успешно создан для 3D документа!",
                                        "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    MessageBox.Show("Не удалось создать атрибут 3D документа.",
                                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                                valueArr.ksDeleteArray();
                            }

                            fVisibl.ksDeleteArray();
                            colKeys.ksDeleteArray();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Не удалось создать тип атрибута.",
                            "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    arr.ksDeleteArray();
                }
            }
        }

        private void BtnShowAttribute_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ConnectToKompas())
                {
                    MessageBox.Show("KOMPAS-3D не запущен или документ не открыт.\nЗапустите KOMPAS-3D и откройте 3D документ (деталь или сборку).",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ShowAttributeValue();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при чтении атрибута:\n{0}", ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowAttributeValue()
        {
            // Получаем ссылку на 3D документ
            reference docRef = doc3D.reference;

            // Создаем параметры и массив ДО цикла (аналогично Step8.cs:470-488)
            ksUserParam usPar = (ksUserParam)kompas.GetParamStruct((short)StructType2DEnum.ko_UserParam);
            ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)StructType2DEnum.ko_LtVariant);
            ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(23);

            if (usPar == null || item == null || arr == null)
            {
                MessageBox.Show("Ошибка инициализации параметров.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            usPar.Init();
            usPar.SetUserArray(arr);

            // Заполняем массив начальными значениями (по аналогии со Step8.cs:479-487)
            item.Init();
            item.strVal = string.Empty;
            arr.ksAddArrayItem(-1, item);

            // Создаем итератор атрибутов для 3D документа
            ksIterator iter = (ksIterator)kompas.GetIterator();
            if (iter != null && iter.ksCreateAttrIterator(docRef, 0, 0, 0, 0, 0))
            {
                // Получаем первый атрибут
                reference pAttr = iter.ksMoveAttrIterator("F", ref docRef);

                bool found = false;

                while (pAttr != 0)
                {
                    // Получаем информацию о типе атрибута
                    int k1 = 0, k2 = 0, k3 = 0, k4 = 0;
                    double attrTypeNum = 0;
                    attr.ksGetAttrKeysInfo(pAttr, out k1, out k2, out k3, out k4, out attrTypeNum);

                    // Проверяем тип атрибута по ключам
                    if (k1 == 100 && k2 == 200 && k3 == 300)
                    {
                        // Читаем строку атрибута (аналогично Step8.cs:526)
                        attr.ksGetAttrRow(pAttr, 0, 0, 0, usPar);

                        // Получаем массив ПОСЛЕ чтения (аналогично Step8.cs:530)
                        ksDynamicArray readArr = (ksDynamicArray)usPar.GetUserArray();

                        if (readArr != null && readArr.ksGetArrayCount() > 0)
                        {
                            // Получаем значение из первой колонки
                            item.Init();
                            if (readArr.ksGetArrayItem(0, item) == 1)
                            {
                                string value = item.strVal;
                                lblAttributeValue.Text = string.Format("Значение: {0}", value);
                                found = true;
                            }
                        }

                        break;
                    }

                    // Переходим к следующему атрибуту
                    pAttr = iter.ksMoveAttrIterator("N", ref docRef);
                }

                if (!found)
                {
                    lblAttributeValue.Text = "Значение: атрибут не найден";
                    MessageBox.Show("Атрибут 'Наименование' не найден в 3D документе.",
                        "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                lblAttributeValue.Text = "Значение: ошибка чтения";
                MessageBox.Show("Не удалось создать итератор атрибутов.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Очищаем массив
            if (arr != null)
                arr.ksDeleteArray();
        }

        private void DeleteDocumentAttribute()
        {
            if (createdAttrTypeNumber == 0)
            {
                MessageBox.Show("Атрибут не был создан или уже удален.\nСначала создайте атрибут с помощью первой кнопки.",
                    "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Получаем ссылку на 3D документ
            reference docRef = doc3D.reference;

            // Создаем итератор атрибутов для 3D документа
            ksIterator iter = (ksIterator)kompas.GetIterator();
            if (iter != null && iter.ksCreateAttrIterator(docRef, 0, 0, 0, 0, 0))
            {
                // Получаем первый атрибут
                reference pAttr = iter.ksMoveAttrIterator("F", ref docRef);

                bool found = false;
                while (pAttr != 0)
                {
                    // Проверяем тип атрибута
                    int k1 = 0, k2 = 0, k3 = 0, k4 = 0;
                    double attrTypeNum = 0;
                    attr.ksGetAttrKeysInfo(pAttr, out k1, out k2, out k3, out k4, out attrTypeNum);

                    // Проверяем, является ли это наш атрибут по ключам
                    if (k1 == 100 && k2 == 200 && k3 == 300)
                    {
                        // Удаляем атрибут
                        int result = attr.ksDeleteAttr(docRef, pAttr, string.Empty);

                        if (result == 1)
                        {
                            MessageBox.Show("Атрибут 'Наименование' успешно удален из 3D документа!",
                                "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            found = true;
                            createdAttrTypeNumber = 0;
                            break;
                        }
                        else
                        {
                            MessageBox.Show("Не удалось удалить атрибут.",
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    // Переходим к следующему атрибуту
                    pAttr = iter.ksMoveAttrIterator("N", ref docRef);
                }

                if (!found)
                {
                    MessageBox.Show("Атрибут 'Наименование' не найден в 3D документе.",
                        "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Не удалось создать итератор атрибутов.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
