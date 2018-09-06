using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Data;

namespace ExcelImport
{
    class Program
    {
        private static string _url = "http://rvspapent"; 

        static string workbookId;
        static DataTable dt = new DataTable();
        static List<Column> columnList = new List<Column>();
        static int startColumn;
        static int endColumn;
        static int startRow;
        static int endRow;
        static string pathToExcel;
        static string url;
        static string listName;


       

        static void ConsoleWrite(string message, ConsoleColor color)
        {
            ConsoleColor c = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ForegroundColor = c;
        }


        static void CreateSPListItems(SPList list)
        {
            int i = 1;
            foreach (DataRow row in dt.Rows)
            {
                SPListItem item = list.AddItem();
                i++;
                foreach (Column column in columnList)
                {
                   

                    if (!row.IsNull(column.Name))
                    {
                        string v = row[column.Name].ToString();

                        if (column.FieldType == Microsoft.SharePoint.SPFieldType.User)
                        {
                            SPUser user = list.ParentWeb.EnsureUser(v);
                            item[column.Name] = user;
                        }
                        else if (column.FieldType == Microsoft.SharePoint.SPFieldType.MultiChoice)
                        {
                            string[] singleValues = v.Split(';');
                            var choicevalues = new SPFieldMultiChoiceValue();
                            foreach (var s in singleValues)
                            {
                                choicevalues.Add(s);
                            }

                            item[column.Name] = choicevalues;
                        }
                        else if (column.FieldType == Microsoft.SharePoint.SPFieldType.Boolean)
                        {
                            v = v.ToLower();
                            bool? newValue = null;
                            if (v == "falsch" || v == "nein" || v == "0" || v == "no" || v == "false")
                            {
                                newValue = false;
                            }
                            else if (v == "ja" || v == "wahr" || v == "1" || v == "yes" || v == "true")
                            {
                                newValue = true;
                            }

                            if (newValue.HasValue)
                            {
                                item[column.Name] = newValue.Value;
                            }
                        }
                        else if (column.FieldType == Microsoft.SharePoint.SPFieldType.DateTime)
                        {
                            DateTime dateTime;
                            if(DateTime.TryParse(v, out dateTime))
                            {
                                item[column.Name] = dateTime;
                            }
                            else
                            {
                                throw new Exception(string.Format("Can not convert value [{0}] in column [{1}] to Double (Row: {2})", v, column.Name, i));
                            }
                            
                        }
                        else if (column.FieldType == Microsoft.SharePoint.SPFieldType.Number)
                        {
                            double d;
                            if (double.TryParse(v,out d))
                            {
                                item[column.Name] = d;
                            }
                            else
                            {
                                throw new Exception(string.Format("Can not convert value [{0}] in column [{1}] to Double (Row: {2})", v,column.Name, i));
                            }
                        }
                        else if (column.FieldType == Microsoft.SharePoint.SPFieldType.Lookup)
                        {
                            var lookupField = list.Fields[column.Name] as SPFieldLookup;
                            
                            var lookupList = list.ParentWeb.Lists[new Guid(lookupField.LookupList)];
                            SPQuery query = new SPQuery();
                            query.RowLimit = 1;                            
                            query.Query = "<Where><Eq><FieldRef Name='" + lookupField.LookupField + "' /><Value Type='Text'>" + v + "</Value></Eq></Where>";
                            query.ViewFieldsOnly = true;
                            query.ViewFields = "<FieldRef Name='ID' />";
                            SPListItemCollection itemCollection = lookupList.GetItems(query);
                            if (itemCollection.Count > 0)
                            {
                                var lookupitem = itemCollection[0];
                                var lookupValue = new SPFieldLookupValue(lookupitem.ID, lookupitem.ID.ToString());
                                item[column.Name] = lookupValue;
                            }                                                        
                        }
                        else
                        {
                            item[column.Name] = v;
                        }
                    }                   
                }                
                item.Update();
                ConsoleWrite(string.Format("Create Item {0} of {1}", i, dt.Rows.Count -1), ConsoleColor.Green);
            }

        }

        static void Main(string[] args)
        {
            ConsoleWrite("This application imports an Excel file into SharePoint.", ConsoleColor.Blue);
            ConsoleWrite("The Excel file must have column names in the first line!", ConsoleColor.Blue);
            Console.Write("SharePoint Website Url: ");
            url = Console.ReadLine();
            Console.Write("SharePoint Listname: ");
            listName = Console.ReadLine();
            Console.Write("Path to Excel file: ");
            pathToExcel = Console.ReadLine(); 
           
            try
            {
                Start();
            }
            catch (Exception ex)
            {
                ConsoleWrite(ex.Message, ConsoleColor.Red);
            }
            finally
            {
                ConsoleWrite("End", ConsoleColor.DarkGray);
                Console.ReadLine();
            }
        }

        private static void Start()
        {
            using (SPSite site = new SPSite(url))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    ConsoleWrite(string.Format("SharePoint site [{0}] found", web.Title), ConsoleColor.Green);

                    SPList list = web.Lists[listName];
                    ConsoleWrite(string.Format("SharePoint list [{0}] found", list.Title), ConsoleColor.Green);

                    var package = new ExcelPackage(new FileInfo(pathToExcel));

                    if (package.Workbook.Worksheets.Count > 0)
                    {
                        foreach (ExcelWorksheet sheet in package.Workbook.Worksheets)
                        {
                            ConsoleWrite(string.Format("Found Worksheet {0} with Id: {1}", sheet.Name, sheet.Index), ConsoleColor.Green);
                        }

                        Console.Write("Select WorksheetId: ");
                        workbookId = Console.ReadLine();

                    }
                    else
                    {
                        throw new Exception("No Worksheets found in file!");
                    }

                    ExcelWorksheet workSheet = package.Workbook.Worksheets[int.Parse(workbookId)];
                    startColumn = workSheet.Dimension.Start.Column;
                    endColumn = workSheet.Dimension.End.Column;
                    startRow = workSheet.Dimension.Start.Row;
                    endRow = workSheet.Dimension.End.Row;

                    ConsoleWrite(string.Format("Start Row: {0}", startRow), ConsoleColor.White);
                    ConsoleWrite(string.Format("End Row: {0}", endRow), ConsoleColor.White);
                    ConsoleWrite(string.Format("Start Column: {0}", startColumn), ConsoleColor.White);
                    ConsoleWrite(string.Format("End Column: {0}", endColumn), ConsoleColor.White);

                    ConsoleWrite("Start reading the header... ", ConsoleColor.Cyan);
                    List<string> header = new List<string>();
                    for (int i = workSheet.Dimension.Start.Column; i <= workSheet.Dimension.End.Column; i++)
                    {
                        string columnName = workSheet.Cells[1, i].Value.ToString();
                        // Check if columns exists in SharPoint List
                        if (list.Fields.ContainsField(columnName))
                        {
                            if (columnList.Exists(x => x.Name == columnName))
                            {
                                throw new Exception(string.Format("Column [{0}] exists several times in Excel!", columnName));
                            }

                            SPField field = list.Fields.GetField(columnName);
                            columnList.Add(new Column(columnName, i, field.Type));
                            dt.Columns.Add(columnName);
                            ConsoleWrite(string.Format("Field [{0}] with the type [{1}] found in SharePoint list", columnName, field.Type), ConsoleColor.Green);
                        }
                        else
                        {
                            ConsoleWrite(string.Format("The field [{0}] does not exist in SharePoint list", columnName), ConsoleColor.Yellow);
                        }


                    }
                    ConsoleWrite("Header creation completed", ConsoleColor.White);
                    ConsoleWrite("Start reading row values...", ConsoleColor.White);

                    for (int i = startRow + 1; i <= endRow; i++)
                    {
                        ConsoleWrite(string.Format("Current Row: {0}", i), ConsoleColor.White);
                        DataRow row = dt.NewRow();
                        bool rowIsNotEmpty = false;

                        foreach (Column c in columnList)
                        {
                            if (workSheet.Cells[i, c.Index].Value == null)
                            {
                                row[c.Name] = DBNull.Value;
                            }
                            else
                            {
                                row[c.Name] = workSheet.Cells[i, c.Index].Value;
                                rowIsNotEmpty = true;
                            }
                        }

                        if (rowIsNotEmpty)
                        {
                            dt.Rows.Add(row);
                        }

                    }

                    CreateSPListItems(list);
                }
            }
        }
    }
}
