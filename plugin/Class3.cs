/*
using System;
using System.Collections.Generic;
using System.Linq;

using Autodesk.Civil.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Aec.PropertyData.DatabaseServices;
using Autodesk.AutoCAD.Colors;
//using Autodesk.AutoCAD.Windows;
//using Autodesk.AutoCAD.EditorInput;
using ExcelDataReader;
using System.IO;
using System.Globalization;
using System.Collections.Specialized;
using Autodesk.Windows;

namespace plugin
{
    public class Class3
    {
        Document acDoc = Application.DocumentManager.MdiActiveDocument; //Autocad Application documento
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor; // Autocad EditorInput
        CivilDocument doc = CivilApplication.ActiveDocument; //Civil Application Document
        Database acCurDb = Application.DocumentManager.MdiActiveDocument.Database; //Autocad Database

        // Create propertySet
        public void ChangePropertyName(List<String> columnNames, string name)
        {
            var propertySetDefinition = new PropertySetDefinition(); // Instanciando classe propertySet
            propertySetDefinition.SetToStandard(acCurDb);
            propertySetDefinition.SubSetDatabaseDefaults(acCurDb);
            StringCollection appliesTo = new StringCollection(); // Instanciando StringColection
            propertySetDefinition.AlternateName = name; // Name propertySet
            propertySetDefinition.IsLocked = false;
            propertySetDefinition.IsVisible = true;
            propertySetDefinition.IsWriteable = true;
            appliesTo.Add("AcDb3dSolid"); //Info o objeto que no appliesTo vai ser usado
            propertySetDefinition.SetAppliesToFilter(appliesTo, false); //Adicionando no propertySet

            PropertyDefinition propertyDefinition; //Tabela propertyDefinition
            for (int i = 0; i < columnNames.Count; i++)
            {
                propertyDefinition = new PropertyDefinition();
                propertyDefinition.SetToStandard(acCurDb);
                propertyDefinition.SubSetDatabaseDefaults(acCurDb);
                propertyDefinition.Name = columnNames[i]; //Nome
                propertyDefinition.DataType = Autodesk.Aec.PropertyData.DataType.Text; //Tipo
                propertySetDefinition.Definitions.Add(propertyDefinition);
                propertySetDefinition.SetDisplayOrder(propertyDefinition, i + 1);
            }

            using (Transaction tr = acCurDb.TransactionManager.StartTransaction())
            {
                var dictionaryPropertySetDefinitions = new DictionaryPropertySetDefinitions(acCurDb);
                if (dictionaryPropertySetDefinitions.Has(name, tr))
                {
                    return;
                }

                dictionaryPropertySetDefinitions.AddNewRecord(name, propertySetDefinition); // criando dicionário de property set
                tr.AddNewlyCreatedDBObject(propertySetDefinition, true);

                tr.Commit();
            }
        }
        // Get propertySet
        public static Autodesk.AutoCAD.DatabaseServices.ObjectId GetPropertySetDefinitionIdByName(string psdName)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            Autodesk.AutoCAD.DatabaseServices.ObjectId psdId = Autodesk.AutoCAD.DatabaseServices.ObjectId.Null;
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = db.TransactionManager;

            using (Transaction tr = tm.StartTransaction())
            {
                try
                {
                    DictionaryPropertySetDefinitions psdDict = new DictionaryPropertySetDefinitions(db);
                    if (psdDict.Has(psdName, tr))
                    {
                        psdId = psdDict.GetAt(psdName); // Retorna um dicionário do set de propriedades
                    }
                }
                catch
                {
                    ed.WriteMessage("\n GetPropertySetDefinitionIdByName failed");
                }
                tr.Commit();
                return psdId;
            }
        }
        // PropertySet to solid 
        public static bool AddStairPropertySetToSolid(Solid3d sol, string name)
        {
            bool result = false;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            Autodesk.AutoCAD.DatabaseServices.ObjectId psdId = GetPropertySetDefinitionIdByName(name); // pegando o id do set de propriedades pelo nome

            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                try
                {
                    Autodesk.AutoCAD.DatabaseServices.DBObject dbobj = tr.GetObject(sol.Id, OpenMode.ForWrite); // Pegando objecto do cylinder
                    PropertyDataServices.AddPropertySet(dbobj, psdId); // adicionando o objecto cylinder ao set de propriedades
                    result = true;
                }
                catch
                {
                    result = false;
                    ed.WriteMessage($"\n AddStairPropertySetToSolid function failed {psdId}");
                }
                tr.Commit();
                return result;
            }
        }
        // Create cylinder
        public static bool SetStairPropertiesToSolid(Solid3d sol, string psdName, List<string> columnNames, List<Object> lista_property_add)
        {
            bool result = false;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection setIds = PropertyDataServices.GetPropertySets(sol);
                if (setIds.Count > 0)
                {
                    foreach (Autodesk.AutoCAD.DatabaseServices.ObjectId id in setIds)
                    {
                        PropertySet pset = (PropertySet)id.GetObject(OpenMode.ForWrite);
                        if (pset.PropertySetDefinitionName == psdName && pset.IsWriteEnabled)
                        {
                            for (int i = 0; i < lista_property_add.Count; i++)
                            {
                                pset.SetAt(pset.PropertyNameToId(columnNames[i]), lista_property_add[i].ToString()); // Pegar pelo id do nome do set propriedades, e adicionar o item
                            }
                            result = true;
                            break;
                        }
                    }
                }
                tr.Commit();
            }
            return result;
        }

        public void CreateCylinder(double N, double E, double Z, double radius, double height, string name, List<string> columnNames, List<Object> lista_property_add, int index)
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
                cylinder.SetDatabaseDefaults();
                cylinder.Color = Color.FromColorIndex(ColorMethod.ByLayer, (short)index);
                cylinder.CreateExtrudedSolid(circle, extrusionVector, new SweepOptions());

                // Add the Solid to the model space
                BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // Add the cylinder to the current space
                currentSpace.AppendEntity(cylinder);
                tr.AddNewlyCreatedDBObject(cylinder, true);

                // Remove the circle from the current space
                circle.Erase();

                // Add propertySet to solid
                AddStairPropertySetToSolid(cylinder, name);
                // Add item in propertyset to solid
                SetStairPropertiesToSolid(cylinder, name, columnNames, lista_property_add);
                // Commit the transaction
                tr.Commit();
            }
        }

        public void readExcel(string filepath, int diametro, string nameSet)
        {
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

                    int indice_nspt = ColumnNames.FindIndex(x => x.StartsWith("NSPT")); // primeiro nspt
                    int max_nspt = ColumnNames.FindLastIndex(item => item.StartsWith("NSPT")); // último nspt
                    int indice_camada = ColumnNames.FindIndex(x => x.StartsWith("CAM"));
                    int max_camada = ColumnNames.FindLastIndex(item => item.StartsWith("CAM"));
                    List<string> nomesNSPT = ColumnNames.GetRange(indice_nspt, (max_nspt + 1) - indice_nspt); //ColumnNames.Count - indice_nspt
                    nomesNSPT.Insert(0, ColumnNames[0]);
                    nomesNSPT.Insert(1, ColumnNames[4]);
                    nomesNSPT.Insert(2, "CAM");
                    nomesNSPT.Insert(3, "Cota inicial");
                    nomesNSPT.Insert(4, "Cota final");
                    nomesNSPT.Insert(5, "Cota Terreno");
                    nomesNSPT.Insert(6, "ESPESSURA");

                    // Create PropertySet, e add itens in propertySetDefinition
                    ChangePropertyName(nomesNSPT, nameSet);

                    foreach (var row in rowData)
                    {
                        var lista_nspt = new List<object> { row[0] }; //Código e NA(m)
                        List<object> camada = new List<object>();
                        List<object> list = new List<object> { row[1], row[2], row[3], row[4] };
                        for (int index = indice_camada; index < max_camada + 1; index += 2)
                        {
                            if (row[index] == null || row[index] == "")
                            {
                                break;
                            }
                            else
                            {
                                string tipo_areia = row[index].ToString();
                                string espessura = row[index + 1].ToString();
                                string[] array_espessura = espessura.Split(new string[] { " a " }, StringSplitOptions.None);
                                array_espessura = array_espessura.Concat(new string[] { tipo_areia }).ToArray();
                                camada.Add(array_espessura);
                            }
                        }
                        for (int i = indice_nspt; i <= (max_nspt + 1) - indice_nspt; i++)
                        {
                            if (row[i] == null || row[i] == "") break;
                            lista_nspt.Add(row[i]);
                        }
                        list_data.Add(new List<object> { list, camada, lista_nspt });
                        cont++;
                    }

                    // Passando por todas as coordenadas e camadas

                    for (var indice = 0; indice < rowData.Count; indice++)
                    {
                        // ------- Array 1 - XYZ coordenadas --------
                        var listaDouble2 = ((List<object>)((List<object>)list_data[indice])[0]); // Lista array 1
                        if (listaDouble2[0] == "" || listaDouble2[0] == null) break; // se o N for igual a vazio, significa que é a última linha
                        double N = Convert.ToDouble(listaDouble2[0]); // X
                        double E = Convert.ToDouble(listaDouble2[1]); // Y
                        double Z = Convert.ToDouble(listaDouble2[2]);
                        string NA = Convert.ToString(listaDouble2[3]); // Nível da Água

                        // ------- Array 2 - Camadas --------
                        var qtd_camadas = ((List<object>)((List<object>)list_data[indice])[1]).Count; // Quantidade de Camadas           

                        for (int index = 0; index < qtd_camadas; index++)
                        {
                            var lista_property_add = new List<object>(((List<object>)((List<object>)list_data[indice])[2]));
                            lista_property_add.Insert(1, NA.ToString()); //examplo
                            var test_layer = ((string[])((List<object>)((List<object>)list_data[indice])[1])[index]); // Lista array 2
                            string tipo_areia = test_layer[2]; // Tipo de areia
                            string espessura_ini = (test_layer[0]).Replace(',', '.');
                            string espessura_fim = (test_layer[1]).Replace(',', '.');
                            double ini_value = Double.Parse(espessura_ini); // Valor 1 de espessura camada
                            double fim_value = Double.Parse(espessura_fim); // Valor 2 de espessura camada
                            lista_property_add.Insert(2, tipo_areia);
                            lista_property_add.Insert(3, espessura_ini);
                            lista_property_add.Insert(4, espessura_fim);
                            lista_property_add.Insert(5, Z.ToString());
                            lista_property_add.Insert(6, (fim_value - ini_value).ToString());

                            if (ini_value == 0) //subtraindo a espessura inicial - Z
                            {
                                CreateCylinder(E, N, Z, diametro, fim_value * (-1), nameSet, nomesNSPT, lista_property_add, index); // chamar função [N,E,Z,NA,height]    
                            }
                            else
                            {
                                double height = (fim_value - ini_value) * (-1); // espessura
                                CreateCylinder(E, N, Z - ini_value, diametro, height, nameSet, nomesNSPT, lista_property_add, index); // chamar função
                            }
                        }
                    }
                }
            }
        }

        [CommandMethod("cilindros_de_sondagem")]
        public void Form()
        {
            Form1 form = new Form1(); // Iniciar Form 1
            form.ShowDialog();
        }

        private static Autodesk.Windows.RibbonControl ribbon;
        private static RibbonTab lateralTab;

        [CommandMethod("Menu_teste")]
        public void Initialize()
        {
            ribbon = ComponentManager.Ribbon;

            // Create the lateral menu tab
            lateralTab = new RibbonTab();
            lateralTab.Title = "TPF Tools";
            ribbon.Tabs.Add(lateralTab);

            // Create a panel for the lateral menu content
            RibbonPanel panel = new RibbonPanel();
            lateralTab.Panels.Add(panel);

        }

    }
}
*/