using System;
using System.Collections.Generic;
using System.Data;
using Excel.Models;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;



namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            #region 

            
            //Veri Ekleme İşlemi
            using (var context = new Context())
            {
                #region
                // ekleyip yazdırma
                //var bilgi = new Bilgiler()
                //{
                //    Ad = "Şafak",
                //    Soyad = "Özdemir",
                //    Sinif = "4",
                //    Yas = 24

                //};

                //context.Bilgilers.Add(bilgi);
                //context.SaveChanges();
                ///////////////////////////////////////
                #endregion
                Context c = new Context();
                var a = c.Bilgilers; // 

                #region
                foreach (var item in a)
                {
                    Console.WriteLine(item.Ad);
                }


                object json = JsonConvert.SerializeObject(a); 
                




                List<Bilgiler> List = JsonConvert.DeserializeObject<List<Bilgiler>>(json.ToString());


                DataTable table = Helpers.ToDataTable<Bilgiler>(List);  

                

               
                using (SpreadsheetDocument document = SpreadsheetDocument.Create("D:/TestNewData.xlsx", SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                    sheets.Append(sheet);

                    Row headerRow = new Row();
                    /// Ne kadar sütun varsa ekliyo aşağısı
                    List<String> columns = new List<string>();
                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }
                    //////////
                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        Row newRow = new Row();
                        foreach (String col in columns)
                        {
                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(dsrow[col].ToString());
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }

                    workbookPart.Workbook.Save();
                }
                #endregion


            }

            #endregion

           





        }




















    }















}


