using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace GroupWorkSubmission1
{
    class Program
    {
        static Excel.Workbook workbook;
        static Excel.Application app;

        static void Main(string[] args)
        {
            app = new Excel.Application();
            app.Visible = true;
            try
            {
                workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
                workbook.Save();
                workbook.Close();
                app.Quit();
            }
            catch
            {
                SetUp();
                app.Quit();
            }

            var input = "";
            while (input != "x")
            {
                PrintMenu();
                input = Console.ReadLine();
                try
                {
                    var option = int.Parse(input);
                    switch (option)
                    {
                        case 1:
                            try
                            {
                                Console.Write("Enter the size: ");
                                var size = float.Parse(Console.ReadLine());
                                Console.Write("Enter the suburb: ");
                                var suburb = Console.ReadLine();
                                Console.Write("Enter the city: ");
                                var city = Console.ReadLine();
                                Console.Write("Enter the market value: ");
                                var value = float.Parse(Console.ReadLine());

                                AddPropertyToWorksheet(size, suburb, city, value);
                            }
                            catch
                            {
                                Console.WriteLine("Error: couldn't parse input");
                            }
                            break;
                        case 2:
                            Console.WriteLine("Mean price: " + CalculateMean());
                            break;
                        case 3:
                            Console.WriteLine("Price variance: " + CalculateVariance());
                            break;
                        case 4:
                            Console.WriteLine("Minimum price: " + CalculateMinimum());
                            break;
                        case 5:
                            Console.WriteLine("Maximum price: " + CalculateMaximum());
                            break;
                        default:
                            break;
                    }
                }
                catch { }
            }
        }

        static void PrintMenu()
        {
            Console.WriteLine();
            Console.WriteLine("Select an option (1, 2, 3, 4, 5) " +
                              "or enter 'x' to quit...");
            Console.WriteLine("1: Add Property");
            Console.WriteLine("2: Calculate Mean");
            Console.WriteLine("3: Calculate Variance");
            Console.WriteLine("4: Calculate Minimum");
            Console.WriteLine("5: Calculate Maximum");
            Console.WriteLine();
        }

        static void SetUp()
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook wb = app.Workbooks.Add();

            var awb = app.ActiveWorkbook;

            Excel.Worksheet currentsheet = (Excel.Worksheet)wb.Sheets[1];
            currentsheet.Name = "Property Details";
            currentsheet.Cells[1, "A"] = "Size (in square feet)";
            currentsheet.Cells[1, "B"] = "Suburb";
            currentsheet.Cells[1, "C"] = "City";
            currentsheet.Cells[1, "D"] = "Market Value";

            awb.SaveAs2("property_pricing.xlsx");

            // save before exiting
            awb.Save();
            awb.Close();
            app.Quit();

        }

        static void AddPropertyToWorksheet(float size, string suburb, string city, float value)
        {
            int row = 2;
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
            Excel.Worksheet propertysheet = (Excel.Worksheet)workbook.Sheets[1];

            while (true)
            {
                // look for first empty row and enter property details there
                if (propertysheet.Range["A" + row].Value == null)
                {
                    propertysheet.Cells[row, "A"] = size;
                    propertysheet.Cells[row, "B"] = suburb;
                    propertysheet.Cells[row, "C"] = city;
                    propertysheet.Cells[row, "D"] = value;

                    // set row-counter in a known cell
                    propertysheet.Cells[1, "E"] = row; // row-counter stored to the right of the last header

                    workbook.Save();
                    workbook.Close();
                    app.Quit();
                    return;
                }
                row++;
            }

        }

        static float CalculateMean()
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
            Excel.Worksheet propertysheet = (Excel.Worksheet)workbook.Sheets[1];

            var rowcount = Convert.ToInt32(propertysheet.Range["E1"].Value);

            float sum = 0.0f;
            for (int i = 2; i<= rowcount; i++)
            {
                sum = (sum + Convert.ToSingle(propertysheet.Range["D" + i].Value));
            }

            app.Quit();
            return sum/ (rowcount-1); // subtract 1 because actual count = rowcount - 1 as headers exist

        }

        static float CalculateVariance()
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
            Excel.Worksheet propertysheet = (Excel.Worksheet)workbook.Sheets[1];

            var rowcount = Convert.ToInt32(propertysheet.Range["E1"].Value);

            float mean = CalculateMean();
            float sqsum = 0.0f;
            for (int i = 2; i<= rowcount; i++)
            {
                // calculating sum of squared differences from the mean
                sqsum = sqsum + ((Convert.ToSingle(propertysheet.Range["D" + i].Value) - mean) * (Convert.ToSingle(propertysheet.Range["D" + i].Value) - mean));
            }

            app.Quit();
            return sqsum/(rowcount - 2); // subtract 1 because actual count = rowcount - 1 as headers exist
                                         // subtract another 1 because we are calculating sample variance 
        }

        static float CalculateMinimum()
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
            Excel.Worksheet propertysheet = (Excel.Worksheet)workbook.Sheets[1];

            var rowcount = Convert.ToInt32(propertysheet.Range["E1"].Value);

            float min = Convert.ToSingle(propertysheet.Range["D2"].Value); // initially, min set to 1st value
            for (int i = 3; i <= rowcount; i++) // iterating over next rows
            {
                if (Convert.ToSingle(propertysheet.Range["D" + i].Value) < min)
                {
                    min = Convert.ToSingle(propertysheet.Range["D" + i].Value); // if a lower value is found, min is reset
                }
            }

            app.Quit();
            return min;
        }

        static float CalculateMaximum()
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
            Excel.Worksheet propertysheet = (Excel.Worksheet)workbook.Sheets[1];

            var rowcount = Convert.ToInt32(propertysheet.Range["E1"].Value);

            float max = Convert.ToSingle(propertysheet.Range["D2"].Value); // initially, max set to 1st value
            for (int i = 3; i <= rowcount; i++) // iterating over next rows
            {
                if (Convert.ToSingle(propertysheet.Range["D" + i].Value) > max)
                {
                    max = Convert.ToSingle(propertysheet.Range["D" + i].Value); // if a higher value is found, max is reset
                }
            }

            app.Quit();
            return max;

        }
    }
}
