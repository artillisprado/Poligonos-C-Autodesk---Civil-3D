using ExcelDataReader;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


readExcel();
void readExcel()
{
    var filepath = "C:\\Users\\artillis.prado\\Downloads\\NSPT solido.xlsx";
    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    CultureInfo.CurrentCulture = new CultureInfo("en-US");
    using (var stream = File.Open(filepath, FileMode.Open, FileAccess.Read))
    {
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {

            // Read the data from the sheet
            DataSet result = reader.AsDataSet();

            // Access the data from the first table in the dataset
            DataTable table = result.Tables[0];

            // Max rows and max columns
            var row_count = 0;
            var column_count = 0;
 
            if (result.Tables.Count > 0)
            {
                row_count = table.Rows.Count;
                column_count = table.Columns.Count;
            }

            // Create a list to hold the data
            List<List<object>> rowData = new List<List<object>>();

            // Create a list to hold the values for the current row
            List<string> ColumnNames = new List<string>();

            int contador = 0;
            // Read the data rows
            while (reader.Read())
            {
                if (contador == row_count)
                {
                    break;
                }
                // Create a list to hold the values for the current row
                List<object> rowValues = new List<object>();

                if (contador > 0)
                {
                    // Read the values for each column in the current row
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        // Add the value to the current row list
                        rowValues.Add(reader.GetValue(i));
                    }

                    // Add the row list to the main list of row data
                    rowData.Add(rowValues);

                }
                else
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        // Add the value to the current row list
                        ColumnNames.Add(reader.GetValue(i).ToString());
                    }
                }
                contador++;
            }

            int cont = 0;
            // lista de dados do excel [[x,y,z][camada,espessura]]
            IList<object> list_data = new List<object>();
            // nomes de colunas do set property
            int indice_nspt = ColumnNames.FindIndex(x => x == "NSPT_1m-2m");
            int indice_camada = ColumnNames.FindIndex(x => x == "CAM1");
            List<string> namseNSPT = ColumnNames.GetRange(indice_nspt, ColumnNames.Count - indice_nspt);
            namseNSPT.Insert(0, ColumnNames[0]);
            namseNSPT.Insert(1, ColumnNames[4]);
            namseNSPT.Insert(2, "CAM");
            namseNSPT.Insert(3, "ESPESSURA");

            foreach (var row in rowData)
            {
                var lista_nspt = new List<object> { row[0] }; //Código e NA(m)
                List<object> camada = new List<object>();
                List<object> list = new List<object> { row[1], row[2], row[3], row[4] };

                for (int index = indice_camada;index < column_count;index += 2)
                {
                    if (row[index] == null)
                    {
                        break;
                    } else
                    {
                        //lista_nspt.AddRange(new List<Object> { row[index], row[index + 1]}); //Camada e Espessura
                        string tipo_areia = row[index].ToString();
                        string espessura = row[index + 1].ToString();
                        string[] array_espessura = espessura.Split(new string[] { " a " }, StringSplitOptions.None);
                        array_espessura = array_espessura.Concat(new string[] { tipo_areia}).ToArray();
                        camada.Add(array_espessura);
                    }
                }

                for (int i = indice_nspt; i <= ColumnNames.Count - indice_nspt; i++)
                {
                    if (row[i] == null) break;
                    lista_nspt.Add(row[i]);
                }
                list_data.Add(new List<object> { list, camada, lista_nspt });
                cont++;
            }

            // Passando por todas as coordenadas e camadas
            for (var indice = 0; indice < rowData.Count; indice++)
            {
                // ------- Array 1 - XYZ coordenadas --------
                var listaDouble2 = ((List<object>)((List<object>)list_data[indice])[0]).ConvertAll(obj => (double)obj); // Lista array 1
                double N = listaDouble2[0];
                double E = listaDouble2[1];
                double Z = listaDouble2[2];
                double NA = listaDouble2[3];

                // ------- Array 2 - Camadas --------
                var qtd_camadas = ((List<object>)((List<object>)list_data[indice])[1]).Count; // Quantidade de Camadas           
                for (int index = 0; index < qtd_camadas; index++)
                {
                    var test_layer = ((string[])((List<object>)((List<object>)list_data[indice])[1])[index]); // Lista array 2
                    string tipo_areia = test_layer[2]; // Tipo de areia
                    string espessura_ini = (test_layer[0]).Replace(',','.');
                    string espessura_fim = (test_layer[1]).Replace(',', '.');
                    double ini_value = Double.Parse(espessura_ini); // Valor 1 de espessura camada
                    double fim_value = Double.Parse(espessura_fim); // Valor 2 de espessura camada
                    var lista_property_add = new List<object>(((List<object>)((List<object>)list_data[indice])[2]));
                    lista_property_add.Insert(1, NA.ToString());
                    lista_property_add.Insert(2, tipo_areia);
                    lista_property_add.Insert(3, espessura_ini);
                    lista_property_add.Insert(4, espessura_fim);
                    lista_property_add.Insert(5, (fim_value - ini_value).ToString());

                    if (ini_value > 0)
                    {
                        Z = Z - ini_value; // subtraindo altura para próxima posição de inicio
                        double height = Math.Round(fim_value - ini_value,2); // espessura
                    } else
                    {
                        double height = fim_value - ini_value; // espessura
                    }

                    
                }
            }
        }
    }
}



