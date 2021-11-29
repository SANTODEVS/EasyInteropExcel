    public static void Arq_Conec(String nomearq, string nomeSheet, string[] listaCelValidas, int linha)
        {


            // verificando se o arquivo existe no diretorio
            string arqui = Dir_Template + Path.GetFileName(nomearq);

            //abrindo excel
            if (oExcel == null)
            {
                oExcel = new Excel.Application();
            }
            oExcel.Visible = false;
            oExcel.Workbooks.Open(arqui);

            int indiceSheet = ValidaSheet("BILLINGOFFER", oExcel);
            if (indiceSheet != -1)
            {
                Excel.Worksheet ws = (Excel.Worksheet)oExcel.Sheets[indiceSheet];
                Excel.Range usedRange = ws.UsedRange;
               
                var ultimaLinhaUsada = usedRange.Count;
                for (int i = linha; i < ultimaLinhaUsada; i++)
                {
                    bool temDados = false;
                    Excel.Range Range1 = ws.Range[$"B{i}", $"BH{i}"];
                    foreach (Excel.Range a in Range1.Rows.Cells)
                    {
                    
                        string celulaEx = Convert.ToString(a.Address);
                        if (celulaEx.Split(new char[] { '$' }, StringSplitOptions.None)[1] == "B")
                        {
                            if (!string.IsNullOrEmpty(a.Value)) temDados = true;
                        }
                        if (ValidaCelula(celulaEx,listaCelValidas,out bool primeiraCelula, out bool ultimaCelula) && temDados)
                        {

                            string valorEx = Convert.ToString(a.Value) is null ? "" : Convert.ToString(a.Value);
                            using (StreamWriter w = File.AppendText(File_Name))
                            {
                                if (!ultimaCelula) w.Write(valorEx + ";");
                                else w.WriteLine(valorEx);

                            }

                        }
                    }
                }
               
                FechaEx(oExcel);
            }
            else 
            { 
                FechaEx(oExcel);
                throw new Exception("Aba desejada não encontrada.");
            }


        }
    

    static void FechaEx(Excel.Application oExcel)
    {
        oExcel.ActiveWorkbook.Close(false);
        oExcel.Quit();
        oExcel = null;
    }

        static bool ValidaCelula(string celulaEX, string[] listaDelValidas) 
        {
            return ValidaCelula(celulaEX, listaDelValidas,out bool q1, out bool q2);
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
                    if (endereco.Equals("BH")) primeiraCelula = true;
                    if (endereco.Equals("BH")) ultimaCelula = true;
                    return true;
                }

            }
            return false;
        }


   

    static int ValidaSheet(string nomeSheet, Excel.Application oExcel)
    {
        var qtdeSheet = oExcel.ActiveWorkbook.Sheets.Count;



        for (int k = 1; k <= qtdeSheet; k++)
        {
            string no = oExcel.ActiveWorkbook.Sheets[k].Name;
            if (no.ToUpper().Equals(nomeSheet.ToUpper()))
            {

                return k;

            }
        }

        return -1;

    }