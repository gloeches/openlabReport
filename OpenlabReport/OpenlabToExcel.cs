using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using Spire.Xls;

namespace OpenlabReport
{
    public class OpenlabToExcel
    {
        public double resultado;
        public String palabra;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        public bool CreateWorkbook(string source, string destination)
        {
            try
            {
                
                Workbook wbToStream = new Workbook();
                FileStream fileStream = File.OpenRead(source);
                fileStream.Seek(0, SeekOrigin.Begin);
                wbToStream.LoadFromStream(fileStream);
                FileStream file_stream = new FileStream(destination, FileMode.Create);
                wbToStream.SaveToStream(file_stream);
                file_stream.Close();
                fileStream.Close();
                return true;
            }
            catch (Exception e)
            {
                Logger.Info("Fail to copy excel from: " + source + " to: " + destination);
                Logger.Info(e.Message);
                return false;
            }
        }
        public double Doblar(double numero)
        {   
            
            return numero*2;
        }
        public string MyUpper(string s)
        {
            return s.ToUpper();
        }

        public double suma(double[] numeros)
        {
            return numeros[0];
        }

        public string ToExcell(string s)
        {
            try
            {
                Workbook wbToStream = new Workbook();
                FileStream fileStream = File.OpenRead("d:\\data\\sample.xlsx");
                fileStream.Seek(0, SeekOrigin.Begin);
                wbToStream.LoadFromStream(fileStream);
                Worksheet sheet = wbToStream.Worksheets[0];

                sheet.Range["A1"].NumberValue = 321.01;
                //           sheet.Range["A1"].NumberFormat = "0";
                sheet.Range["A2"].Text = s;
                FileStream file_stream = new FileStream("d:\\data\\To_stream.xlsx", FileMode.Create);
                wbToStream.CalculateAllValue();
                sheet.Range.Style.Locked = false;
                sheet.Range["A1:A100"].Style.Locked = true;
                sheet.Protect("3000hanover", SheetProtectionType.All);
                wbToStream.SaveToStream(file_stream);
                file_stream.Close();
                fileStream.Close();
                return "ok: " + s;
            }
            catch (Exception e)
            {
                Logger.Info("Fail filling data into sheet:");
                Logger.Info(e.Message);
                return "Fail";
            }
        }

        public string NumberToExcell(string _sourceFile, string _sheet,string _cell, double _number)
        {
            try
            {
                if (_sourceFile != null || _cell != null)
                {
                    Workbook wb = new Workbook();
                    wb.LoadFromFile(_sourceFile);
                    Worksheet sheet;
                    if (_sheet != null)
                    {
                        sheet = wb.Worksheets[_sheet];
                    }
                    else
                    {
                        sheet = wb.Worksheets[0];
                    }


                    sheet.Range[_cell].NumberValue = _number;
                    wb.CalculateAllValue();
                    sheet.Range.Style.Locked = false;
                    sheet.Range["A1:A100"].Style.Locked = true;
                    sheet.Protect("3000hanover", SheetProtectionType.All); 
                    wb.Save();
                    return _cell + ":" + _number;
                }
                else
                {
                    return "Fail";
                }
            }
            catch (Exception e)
            {
                Logger.Info("Fail filling sheet: " + _sheet + " in " );
                Logger.Info(e.Message);
                return "Fail";
            }
        }


    }


}
