using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using NPOI.SS.UserModel;

namespace MSExcelReader
{
    class ExcelHeaderAttribute : Attribute
    {
        public ExcelHeaderAttribute(String headName)
        {
            m_HeadName = headName;
        }

        private String m_HeadName = String.Empty;
        public String HeadName
        {
            get
            {
                return m_HeadName;
            }
            set
            {
                m_HeadName = value;
            }
        }
    }

    public class ExcelReader
    {
        public ExcelReader(String excelName)
        {
            if (!System.IO.File.Exists(excelName))
            {
                excelName = Environment.CurrentDirectory + excelName;
            }
            m_bIsOpen = OpenExcel(excelName);

            if (!m_bIsOpen)
                throw new Exception("Failed to read the file: " + excelName);
        }

        ~ExcelReader()
        {
            m_WorkBook = null;
            if (m_ExcelStream != null)
            {
                m_ExcelStream.Close();
                m_ExcelStream = null;
            }
        }

        //private Excel.Application m_App = null;
        //private Excel.Workbook m_WorkBook = null;

        System.IO.FileStream m_ExcelStream = null;
        IWorkbook m_WorkBook = null;

        public bool OpenExcel(String excelName)
        {
            try
            {
                m_ExcelStream = System.IO.File.OpenRead(excelName);
                if (excelName.EndsWith(".xls"))
                {
                    m_WorkBook = new NPOI.HSSF.UserModel.HSSFWorkbook(m_ExcelStream);
                }
                else if (excelName.EndsWith(".xlsx"))
                {
                    m_WorkBook = new NPOI.XSSF.UserModel.XSSFWorkbook(m_ExcelStream);
                }

                if (m_WorkBook == null)
                {
                    return false;
                }
                return true;
            }
            catch (System.Exception ex)
            {
				string strError = ex.Message;
                return false;
            }
        }

        private bool m_bIsOpen = false;
        public bool IsOpen
        {
            get
            {
                return m_bIsOpen;
            }
        }

        public void Close()
        {
            try
            {
                m_ExcelStream.Close();
                m_WorkBook = null;
            }
            catch (System.Exception)
            {

            }
        }

        private int GetSheetNameIndex(ISheet sheet, String headerName)
        {
            IRow rowHeader = sheet.GetRow(0);
            for (int i = 0; i < rowHeader.Cells.Count; i++)
            {
                ICell cell = rowHeader.Cells[i];
                if (String.IsNullOrWhiteSpace(cell.StringCellValue))
                {
                    break;
                }
                if (cell.StringCellValue == headerName)
                {
                    return i;
                }
            }
            return -1;
        }

        Dictionary<int, String> GetSheetHeader(String sheetName, Type t)
        {
            try
            {
                if (m_WorkBook == null)
                {
                    return null;
                }

                Dictionary<int, String> headerList = new Dictionary<int, string>();
                ISheet sheet = m_WorkBook.GetSheet(sheetName);
                var properties = t.GetProperties();
                foreach (var propert in properties)
                {
                    var attributes = propert.GetCustomAttributes(typeof(ExcelHeaderAttribute), false);
                    if (attributes != null && attributes.Length > 0)
                    {
                        foreach (var attr in attributes)
                        {
                            if (attr is ExcelHeaderAttribute)
                            {
                                ExcelHeaderAttribute ehAttr = attr as ExcelHeaderAttribute;
                                int nIndex = GetSheetNameIndex(sheet, ehAttr.HeadName);
                                if (nIndex != -1)
                                {
                                    headerList.Add(nIndex, propert.Name);
                                }
                            }
                        }
                    }
                }

                return headerList;
            }
            catch (System.Exception)
            {
                return null;
            }

        }

        private static void SetDataPropertyValue(object objInstance, String propertyName, String strValue)
        {
            if (objInstance == null || String.IsNullOrWhiteSpace(strValue))
            {
                return;
            }
            var property = objInstance.GetType().GetProperty(propertyName);
            if (property == null)
            {
                throw new Exception("Invalid property : " + propertyName);
            }
            if (property.PropertyType == typeof(int))
            {
                property.SetValue(objInstance, int.Parse(strValue), new object[] { });
                return;
            }
            else if (property.PropertyType == typeof(String))
            {
                property.SetValue(objInstance, strValue, new object[] { });
                return;
            }
            else if (property.PropertyType == typeof(uint))
            {
                property.SetValue(objInstance, uint.Parse(strValue), new object[] { });
                return;
            }
            else if (property.PropertyType == typeof(short))
            {
                property.SetValue(objInstance, short.Parse(strValue), new object[] { });
                return;
            }
            else if (property.PropertyType == typeof(Char))
            {
                property.SetValue(objInstance, Char.Parse(strValue), new object[] { });
                return;
            }
            else if (property.PropertyType == typeof(ushort))
            {
                property.SetValue(objInstance, ushort.Parse(strValue), new object[] { });
                return;
            }
            else if (property.PropertyType == typeof(byte))
            {
                property.SetValue(objInstance, byte.Parse(strValue), new object[] { });
                return;
            }
            else if (property.PropertyType == typeof(float))
            {
                property.SetValue(objInstance, float.Parse(strValue), new object[] { });
                return;
            }
            else if (property.PropertyType == typeof(double))
            {
                property.SetValue(objInstance, double.Parse(strValue), new object[] { });
                return;
            }
            else if (property.PropertyType == typeof(bool))
            {
                property.SetValue(objInstance, int.Parse(strValue) != 0, new object[] { });
                return;
            }
            throw new Exception("Invalid property type : " + property.PropertyType.ToString());
        }

        public List<T> GetSheetData<T>(String sheetName) where T : new()
        {
            try
            {
                if (m_WorkBook == null)
                {
                    return null;
                }
                Type t = typeof(T);
                Dictionary<int, String> headerDic = GetSheetHeader(sheetName, t);
                if (headerDic == null || headerDic.Count <= 0)
                {
                    return null;
                }

                ISheet sheet = m_WorkBook.GetSheet(sheetName);

                List<T> sheetData = new List<T>();
                for (int row = 1; row <= sheet.LastRowNum; row++)
                {
                    T data = new T();
                    IRow curRow = sheet.GetRow(row);
                    if(curRow == null)
                    {
                        continue;
                    }

                    ICell cell0 = curRow.Cells[0];
                    if (cell0 == null || (cell0.CellType != CellType.Numeric && cell0.CellType != CellType.String))
                    {
                        break;
                    }
                    foreach (var kv in headerDic)
                    {
						ICell cell = curRow.GetCell(kv.Key);
						String cellValue = String.Empty;
						if (null != cell)
						{
							switch (cell.CellType)
							{
								case CellType.Blank:
									break;
								case CellType.Numeric:
									cellValue = cell.NumericCellValue.ToString();
									break;
								case CellType.String:
									cellValue = cell.StringCellValue;
									break;
								case CellType.Boolean:
									cellValue = cell.BooleanCellValue ? "1" : "0";
									break;
								case CellType.Formula:
									cellValue = cell.NumericCellValue.ToString();
									break;
								default:
									break;
							}
						}
                        try
                        {
                            SetDataPropertyValue(data, kv.Value, cellValue);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.StackTrace);
                            return null;
                        }
                    }
                    sheetData.Add(data);
                }

                return sheetData;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine(ex.StackTrace);
                return null;
            }
        }

		public int GetSheetsCount()
		{
			if (m_WorkBook != null)
			{
				return m_WorkBook.NumberOfSheets;
			}

			return 0;
		}

		public string GetSheetNameByIndex(int nIndex)
		{
			string strSheetName = "";
			if (m_WorkBook != null)
			{
				ISheet isheet = m_WorkBook.GetSheetAt(nIndex);
				if (isheet != null)
				{
					strSheetName = m_WorkBook.GetSheetName(nIndex);
					strSheetName = isheet.SheetName;
				}
			}
			return strSheetName;
		}
    }
}
