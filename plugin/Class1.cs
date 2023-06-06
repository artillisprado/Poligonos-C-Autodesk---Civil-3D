using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Civil.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;

using ExcelDataReader;
using System.Data;
using System.IO;
using Autodesk.Aec.DatabaseServices;
using System.Runtime.CompilerServices;
using System.Globalization;

namespace plugin
{
    public class Class1
    {
        Document acDoc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        CivilDocument doc = CivilApplication.ActiveDocument;
        Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database;

        //[CommandMethod("DrawCylinder")]
        public void DrawCylinder(double N, double E,double Z, double radius, double height)
        {
            // Get the current document and database
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // Start a transaction
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Get the current space (model or paper space)
                BlockTableRecord currentSpace = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                // Create a circle with a center point and radius
                Point3d center = new Point3d(N, E, Z);
                Circle circle = new Circle(center, Vector3d.ZAxis, radius);

                // Add the circle to the current space
                currentSpace.AppendEntity(circle);
                tr.AddNewlyCreatedDBObject(circle, true);

                // Extrude the circle to create a cylinder
                Vector3d extrusionVector = new Vector3d(0, 0, height);
                Solid3d cylinder = new Solid3d();
                cylinder.CreateExtrudedSolid(circle, extrusionVector, new SweepOptions());

                // Add the cylinder to the current space
                currentSpace.AppendEntity(cylinder);
                tr.AddNewlyCreatedDBObject(cylinder, true);

                // Commit the transaction
                tr.Commit();
            }
        }

        [CommandMethod("cilindro")]
        public void readExcel()
        {
            var filepath = "C:\\Users\\artillis.prado\\Downloads\\NSPT solido.xlsx";
            CultureInfo.CurrentCulture = new CultureInfo("en-US");
            using (var stream = File.Open(filepath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    // Read the data from the sheet
                    System.Data.DataSet result = reader.AsDataSet();

                    // Access the data from the first table in the dataset
                    System.Data.DataTable table = result.Tables[0];

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
                        for (int index = 5; index < column_count; index += 2)
                        {
                            if (row[index] == null)
                            {
                                break;
                            }
                            else
                            {
                                string tipo_areia = row[index].ToString();
                                string espessura = row[index + 1].ToString();
                                string[] array_espessura = espessura.Split(new string[] { " a " }, StringSplitOptions.None);
                                //string[] array_espessura = espessura.Split(" a ");
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
                            string espessura_ini = (test_layer[0]).Replace(',', '.');
                            string espessura_fim = (test_layer[1]).Replace(',', '.');
                            double ini_value = Double.Parse(espessura_ini); // Valor 1 de espessura camada
                            double fim_value = Double.Parse(espessura_fim); // Valor 2 de espessura camada

                            if (ini_value < 1)
                            {
                                Z = Z - ini_value; // subtraindo altura para próxima posição de inicio
                                double height = Math.Round(fim_value - ini_value, 2); // espessura
                                DrawCylinder(N, E, Z, NA, height); // chamar função [N,E,Z,NA,height]

                            }
                            else
                            {
                                double height = fim_value - ini_value; // espessura
                                DrawCylinder(N, E, Z, NA, height); // chamar função
                            }
                        }
                        break;
                    }
                }
            }
        }

    }
}
