using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using xlExcel = Microsoft.Office.Interop.Excel;


namespace EasyInteropExcel
{
    public static class OExcel
    {
        public enum XlFileFormat
        {
            xlCSV = 6,
            xlCSVMac = 22,
            xlCSVWindows = 23,
            xlCSVMSDOS = 24,
            xlWorkbookDefault = 51,
        }
        public enum TextFormat
        {
            txt = 1,
            csv = 2
        }
        public static DataTable ConvertToDataTable(object[] array)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(array[0]);
            DataTable dt = CreateDataTable(properties);
            if (array.Length != 0)
            {
                foreach (Object o in array)
                {
                    FillDataTable(properties, dt, o);
                }
            }
            return dt;
        }
        private static DataTable CreateDataTable(PropertyDescriptorCollection properties)
        {
            DataTable dt = new DataTable();

            foreach (PropertyDescriptor pi in properties)
            {
                dt.Columns.Add(pi.Name, pi.PropertyType);
            }
            return dt;
        }
        private static void FillDataTable(PropertyDescriptorCollection properties, DataTable dt, object o)
        {
            DataRow dr = dt.NewRow();
            foreach (PropertyDescriptor pi in properties)
            {
                dr[pi.Name] = pi.GetValue(o);
            }
            dt.Rows.Add(dr);
        }
        public static void ToExcel(IEnumerable<object> Base, string savePath, string fileNameWithoutExt, XlFileFormat formatoXL)
        {
            DataTable dt = ConvertToDataTable(Base.ToArray());
            xlExcel.Application app = new xlExcel.Application();
            xlExcel.Workbook wb = app.Workbooks.Add(xlExcel.XlSheetType.xlWorksheet);
            xlExcel.Worksheet ws = (xlExcel.Worksheet)app.Sheets[1];
            xlExcel.Range usedRange = ws.UsedRange;
            var ultimaLinhaUsada = usedRange.Count;
            app.Visible = true;
            var startingRow = 1;
            int qtdeColumn = dt.Columns.Count;
            int qtdeSheets = 1;
            bool IsCSV = fileNameWithoutExt.ToLower().EndsWith(".csv");
            int MaxLinesExcel = 1048576;

            //gravar cabeçalho
            for (int i = 0; i < qtdeColumn; i++)
            {
                ws.Cells[startingRow, i + 1] = dt.Columns[i].ColumnName;
            }
            foreach (DataRow item in dt.Rows)
            {
                startingRow++;
                //// aqui iremos escrever os dados
                for (int i = 0; i < qtdeColumn; i++)
                {
                    ws.Cells[startingRow, i + 1] = item[i];
                }
                if (startingRow == MaxLinesExcel && !IsCSV)
                {
                    qtdeSheets++;
                    ws = (xlExcel.Worksheet)app.Sheets.Add();
                    MaxLinesExcel += MaxLinesExcel;//dobra ele
                }
            };

            //Salvando arquivo no diretório
            string path = Path.Combine(savePath, fileNameWithoutExt);

            if (File.Exists(path)) File.Delete(path);
            app.DisplayAlerts = false;
            wb.SaveAs(path, formatoXL, Type.Missing, Type.Missing, false, false, xlExcel.XlSaveAsAccessMode.xlNoChange, xlExcel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            Thread.Sleep(10000);
            app.DisplayAlerts = true;
            FechaEx(app);
        }
        public static void ToCSV(IEnumerable<object> Base, string savePath, string fileNameWithoutExt, string delimitador = ";")
        {
            DataTable dt = ConvertToDataTable(Base.ToArray());

            foreach (DataRow linhas in dt.Rows)
            {
                foreach (DataColumn colunas in dt.Columns)
                {


                    using (StreamWriter w = File.AppendText(savePath + "\\" + fileNameWithoutExt))
                    {
                        //w.Write()
                    }
                }
            }
        }

        /// <summary>
        /// Gravar um Excel para Txt ou CSV
        /// </summary>
        /// <param name="File">Nome do arquivo + diretorio</param>
        /// <param name="savePath">Diretorio onde sera salvo o arquivo</param>
        /// <param name="Ext"></param>
        /// <param name="nomeSheet"></param>
        /// <param name="listaCelValidas"></param>
        /// <param name="linha"></param>
        /// <param name="CelulaBase">Esta variavel é utilizada para verificar se sera gravado as colunas restantes</param>
        public static void ExcelToWriteTxt(string File, string savePath, TextFormat Ext, string nomeSheet, string[] listaCelValidas, int linha, string CelulaBase, string Delimitador)
        {
            string FileName = Path.Combine(savePath, Path.GetFileNameWithoutExtension(File) + "." + Ext.ToString());
            xlExcel.Application app = new xlExcel.Application
            {
                Visible = false
            };
            app.Workbooks.Open(File);

            int indiceSheet = ValidaSheet(nomeSheet, app);
            if (indiceSheet != -1)
            {
                xlExcel.Worksheet ws = (xlExcel.Worksheet)app.Sheets[indiceSheet];
                xlExcel.Range usedRange = ws.UsedRange;

                var ultimaLinhaUsada = usedRange.Count;
                for (int i = linha; i < ultimaLinhaUsada; i++)
                {
                    //bool temDados = false;
                    xlExcel.Range Range1 = ws.Range[$"{listaCelValidas[0]}{i}", $"{listaCelValidas[listaCelValidas.Length - 1]}{i}"];
                    foreach (xlExcel.Range a in Range1.Rows.Cells)
                    {
                        string celulaEx = Convert.ToString(a.Address);
                        if (celulaEx.Split(new char[] { '$' }, StringSplitOptions.None)[1] == CelulaBase)
                        {
                            if (a.Value == null || string.IsNullOrEmpty(Convert.ToString(a.Value))) break;
                        }
                        if (ValidaCelula(celulaEx, listaCelValidas, out bool primeiraCelula, out bool ultimaCelula))
                        {

                            string valorEx = Convert.ToString(a.Value) is null ? "" : Convert.ToString(a.Value);
                            using (StreamWriter w = System.IO.File.AppendText(FileName))
                            {
                                if (!ultimaCelula) w.Write(valorEx + Delimitador);
                                else w.WriteLine(valorEx);
                            }

                        }
                    }
                }

                FechaEx(app);
            }
            else
            {
                FechaEx(app);
                throw new Exception("Aba desejada não encontrada.");
            }


        }


        static void FechaEx(xlExcel.Application oExcel)
        {
            oExcel.ActiveWorkbook.Close(false);
            oExcel.Quit();
            foreach (var processo in Process.GetProcessesByName("Excel"))
            {
                if (processo.MainWindowTitle == "") processo.Kill();
            }
        }

        static bool ValidaCelula(string celulaEX, string[] listaDelValidas)
        {
            return ValidaCelula(celulaEX, listaDelValidas, out bool q1, out bool q2);
        }
        static bool ValidaCelula(string celulaEX, string[] listaCelValidas, out bool primeiraCelula, out bool ultimaCelula)
        {
            ultimaCelula = false;
            primeiraCelula = false;
            string celula = celulaEX.Split(new char[] { '$' }, StringSplitOptions.None)[1];


            foreach (var endereco in listaCelValidas)
            {
                if (endereco == celula)
                {
                    if (endereco.Equals(listaCelValidas[0])) primeiraCelula = true;
                    if (endereco.Equals(listaCelValidas[listaCelValidas.Length - 1])) ultimaCelula = true;
                    return true;
                }

            }
            return false;
        }




        static int ValidaSheet(string nomeSheet, xlExcel.Application oExcel)
        {
            var qtdeSheet = oExcel.ActiveWorkbook.Sheets.Count;



            for (int k = 1; k <= qtdeSheet; k++)
            {
                //string no = oExcel.ActiveWorkbook.Sheets[k].Name;
                xlExcel.Worksheet Sheet = (xlExcel.Worksheet)oExcel.ActiveWorkbook.Sheets[k];
                if (Sheet.Name.ToUpper().Equals(nomeSheet.ToUpper()))
                {

                    return k;

                }
            }

            return -1;

        }
    }
}
