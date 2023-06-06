using ExcelDataReader;
using Microsoft.VisualBasic;
using System;
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
                contador++;
            }

            int cont = 0;
            IList<object> list_data = new List<object>();

            foreach (var row in rowData)
            {
                List<object> camada = new List<object>();
                List<object> list = new List<object> { row[1], row[2], row[3], row[4] };
                for (int index = 5;index < column_count;index += 2)
                {
                    if (row[index] == null)
                    {
                        break;
                    } else
                    {
                        string tipo_areia = row[index].ToString();
                        string espessura = row[index + 1].ToString();
                        //string[] array_espessura = espessura.Split(" a ");
                        string[] array_espessura = espessura.Split(new string[] { " a " }, StringSplitOptions.None);
                        array_espessura = array_espessura.Concat(new string[] { tipo_areia }).ToArray();
                        camada.Add(array_espessura);
                    }
                }
                list_data.Add(new List<object> { list, camada });
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
                    Console.WriteLine($"{ini_value} {fim_value}");
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



