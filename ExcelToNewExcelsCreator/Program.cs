using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ExcelToNewExcelsCreator
{

    internal class Program
    {
        static int[,] WitchCellsToChange()
        {
            int[,] cellPostitionXY = {
                { 6, 3  },  //Device_number
                { 7, 11 },  //IP (last octet) - Device_1
                { 7, 13 },  //IP (last octet) - Device_2
                { 6, 23 }   //Date
            };

            return cellPostitionXY;
        }
        static void Main(string[] args)
        {
            //Package licence
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //Paths
            string baseFilesDirectoryPath = GetPathToDirectory_Files();
            string newFilesDirectoryPath = baseFilesDirectoryPath + @"\NewFiles";
            baseFilesDirectoryPath += @"\BaseFiles";

            //Check if files exist
            string baseExcelFile = FindFileWithExtension(baseFilesDirectoryPath, "*.xlsx");
            string[] fileNameAndNumber = CheckFileName(baseExcelFile);

            CreateNewDirectory(newFilesDirectoryPath);
            
            // Ask User
            Console.WriteLine("Write below how many NEW files create");
            int amountOfNewFiles = AskUserAboutAmount();
            if (amountOfNewFiles == 0)
            {
                Console.WriteLine("You chose 0 new files");
                CloseApp();
            }

            Console.WriteLine("Write below how many carriers per day");
            int carriersPerDay = AskUserAboutAmount();
            if (carriersPerDay == 0)
            {
                Console.WriteLine("You chose zero, all files with same date");
            }

            //Machina rusza
            string[] valuesR​eadeFromExcel = ReadValuesFromExcel(baseExcelFile);
            string[] newFilesPath = PreaperNamesForNewFiles(fileNameAndNumber, amountOfNewFiles);
            fileNameAndNumber = null;   //Clean memmory?



            // Czy zwinąć to w jedno i dodać jeszcze czytanie przed?
            PrepareAllData(valuesReadeFromExcel, amountOfNewFiles, carriersPerDay);
            string valuFileNumber = valuesReadeFromExcel[0];
            var a = ChangeNumber(valuFileNumber, amountOfNewFiles); 

            string valueIpAddress_1 = valuesReadeFromExcel[1];
            var b = ChangeIpAddress(valueIpAddress_1, amountOfNewFiles);

            string valueIpAddress_2 = valuesReadeFromExcel[2];
            var c = ChangeIpAddress(valueIpAddress_2, amountOfNewFiles);

            string valueDate = valuesReadeFromExcel[3];
            var d = ChangeDate(valueDate, amountOfNewFiles, carriersPerDay);
            // Czy zwinąć to w jedno i dodać jeszcze czytanie przed?


            // 7. Info: where files are
            Console.WriteLine($"\nFiles saved in folder: {newFilesDirectoryPath}");
            Console.WriteLine("Press any key to close the window");
            Console.ReadKey();

            return;
        }
        enum ColorEnumMessage
        {
            yellow,
            red,
        }
        static void ColoredMesage(ColorEnumMessage color, string message)
        {

            switch (color)
            {
                case ColorEnumMessage.yellow:

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;

                case ColorEnumMessage.red:

                    Console.ForegroundColor = ConsoleColor.Red;
                    break;

                default:

                    Console.ResetColor();
                    break;
            }

            Console.WriteLine(message);
            Console.ResetColor();
        }
        static void CreateNewDirectory(string newDirectoryPath)
        {
            //New directory for files
            try
            {
                if (Directory.Exists(newDirectoryPath))
                {
                    Directory.Delete(newDirectoryPath, true); //true, give permision to delete directory and all content
                    Directory.CreateDirectory(newDirectoryPath);
                }
                else
                {
                    Directory.CreateDirectory(newDirectoryPath);
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                Debug.WriteLine("Catch in CreateNewDirectory -> UnauthorizedAccessException");

                Console.WriteLine("No permition to delete directory:\n" + newDirectoryPath);
                CloseApp();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Can't delete or create new directory");
                Console.WriteLine($"Close all files from directory:\n + {newDirectoryPath}");
                Console.ResetColor();

                CloseApp();
            }
        }
        static string FindFileWithExtension(string directoryPath, string extension)
        {
            string[] path = null;
            try
            {
                path = Directory.GetFiles(directoryPath, extension);
            }
            catch (DirectoryNotFoundException ex)
            {
                Console.WriteLine("Directory not found:\n" + directoryPath);
                CloseApp();
            }
            catch 
            {
                Console.WriteLine("Problems with found file with extension: " + extension);
                CloseApp();
            }
            
            if (path.Length == 0)
            {
                Console.WriteLine("File not found, check directory:\n" + directoryPath);
                CloseApp();
            }
            else
            {
                for (int i = 0; i < path.Length; i++)
                {
                    if ('~' != path[i][0]) //Checking if ther is temporary file"~"
                    {
                        return path[i];
                    }
                }

                Console.WriteLine("Valid file not found, check directory:\n" + directoryPath);
                CloseApp();
            }

            return "";
        }
        static string GetPathToDirectory_Files()
        {
            string path = new DirectoryInfo(".").FullName;
            int ile = path.IndexOf("bin") - 1;
            if (ile < 0)
            {
                Console.WriteLine("Director not fount, please contact with your IT department");
                CloseApp();
            }
            else
            {
                path = path.Substring(0, ile);
                path = path + @"\Files";
            }
            return path;
        }
        static void CloseApp([CallerMemberName] string methodName = "")
        {
            Debug.WriteLine($"!!!There was a problem in method: {methodName}");

            Console.WriteLine("\nPress Enter to close app");
            Console.ReadKey();
            Environment.Exit(0);
        }
        static int AskUserAboutAmount()
        {
            Console.Write("Amount: ");
            try
            {
                return Int32.Parse(Console.ReadLine());
            }
            catch (FormatException)
            {
                ColoredMesage(ColorEnumMessage.yellow, "Error: not valid integer value");
                CloseApp();
            }
            catch
            {
                ColoredMesage(ColorEnumMessage.red, "!!!Error durring reading value from keyboard");
                CloseApp();
            }

            return 0;
        }

        //ułomna
        static void PrepareAllData(string[] valuesReadeFromExcel, int amountOfNewFiles, int carriersPerDay)
        {

        }
        static string[] PreaperNamesForNewFiles(string[] fileNameAndNumber, int amountOfNewFiles)
        {
            int localNumber = 0;
            try
            {
                localNumber = Int32.Parse(fileNameAndNumber[1]);
                if (0 > localNumber)
                {
                    Console.WriteLine("Number will be negative");
                    CloseApp();
                }
            }
            catch
            {
                Debug.WriteLine("Error in ChangeNumber -> Int32.Parse(orginNumber)");
                CloseApp();
            }


            int orginalNameNumberLenght = fileNameAndNumber[1].Length;
            string[] newNameForFiles = new string[amountOfNewFiles];
            for (int i = 0; i < amountOfNewFiles; i++)
            {
                newNameForFiles[i] = (localNumber + (i + 1)).ToString();
                newNameForFiles[i] =
                    new string('0', orginalNameNumberLenght - newNameForFiles[i].Length) +
                    newNameForFiles[i] +
                    fileNameAndNumber[2];
            }

            return newNameForFiles;
        }
        static string[] ChangeIpAddress(string orginIpAddress, int amountOfNewFiles)
        {
            int localIpAddress = 0;
            try
            {
                localIpAddress = Byte.Parse(orginIpAddress);
                if (0 > (255 - (localIpAddress + amountOfNewFiles)))
                {
                    Console.WriteLine("Last occted of IP Addres will be bigger then 255");
                    CloseApp();
                }
            }
            catch
            {
                Debug.WriteLine("Error in ChangeIpAddress -> Byte.Parse(orginIpAddress)");
                CloseApp();
            }

            
            string[] newIpAddress = new string[amountOfNewFiles];
            for (int i=0; i < amountOfNewFiles; i++)
            {
                newIpAddress[i] = (localIpAddress + (i+1)).ToString();
            }

            return newIpAddress;
        }
        static string[] ChangeNumber(string orginNumber, int amountOfNewFiles)
        {
            int localNumber = 0;
            try
            {
                localNumber = Int32.Parse(orginNumber);
                if (0 > localNumber)
                {
                    Console.WriteLine("Number will be negative");
                    CloseApp();
                }
            }
            catch
            {
                Debug.WriteLine("Error in ChangeNumber -> Int32.Parse(orginNumber)");
                CloseApp();
            }


            int orginalNumberLenght = orginNumber.Length;
            string[] newNumber = new string[amountOfNewFiles];
            for (int i = 0; i < amountOfNewFiles; i++)
            {
                newNumber[i] = (localNumber + (i + 1)).ToString();
                newNumber[i] = new string('0', orginalNumberLenght - newNumber[i].Length) + newNumber[i];
            }

            return newNumber;
        }
        static string[] ChangeDate(string orginDate, int amountOfNewFiles, int carriersPerDay)
        {
            DateTime localDate = default; // DateTime.Today;

            try
            {
                DateTime.TryParse(orginDate, out localDate);
            }
            catch
            {
                Debug.WriteLine("Error in ChangeDate -> DateTime.TryParse(orginDate, localDate)");
                CloseApp();
            }

            string[] newDate = new string[amountOfNewFiles];
            for (int i=0; i < amountOfNewFiles; i++)
            {
                if ((i+1) % carriersPerDay == 0)
                {
                    localDate = localDate.AddDays(1);
                }

                if (DayOfWeek.Saturday == localDate.DayOfWeek)
                {
                    localDate = localDate.AddDays(2);
                }
                else if (DayOfWeek.Sunday == localDate.DayOfWeek)   
                {
                    localDate = localDate.AddDays(1);
                }

                newDate[i] = localDate.ToString("d");
            }


            return newDate;
        }


        //checkFileName
        static string [] CheckFileName(string baseExcelFile)
        {
            int amountOfLetters = 2;
            int amountOfNumbers = 3;

            static void CloseWithComment() 
            {
                Console.WriteLine("Check name of base excel file. It shuld start as follow 'CAxxx'");
                CloseApp();
            }

            if (baseExcelFile == "")
            {
                Console.WriteLine("Excel file not found");
                CloseApp();
            }


            try
            {
                string fullFileName =
                    baseExcelFile.Substring(
                    baseExcelFile.LastIndexOf(@"\") + 1);


                string fileLetters = fullFileName.Substring(0, amountOfLetters);
                string fileNumber = fullFileName.Substring(amountOfLetters, amountOfNumbers);
                string fileRestOfName = fullFileName.Substring(amountOfLetters + amountOfNumbers);

                if ("CA" != fileLetters.ToUpper())
                {
                    CloseWithComment();
                }


                try
                {
                    if (0 > Int32.Parse(fileNumber))
                    {
                        CloseWithComment();
                    }
                }
                catch
                {
                    Debug.WriteLine("Error in CheckFileName -> Int32.Parse(fileNumber)");
                    CloseWithComment();
                }

                return new string[] { fileLetters, fileNumber, fileRestOfName };

            }
            catch
            {
                Debug.WriteLine("Error in CheckFileName -> try");
                CloseWithComment();
            }

            return new string[] { };
        }
        static string[] ReadValuesFromExcel(string excelFilesPath)
        {
            Queue<string> valuesFromExcel = new Queue<string>();
            int[,] cellsPositionsXY = WitchCellsToChange();
            int amountOfValueToRead = cellsPositionsXY.Length / 2;

            //open excel file
            var excelWorksheet = new ExcelPackage(excelFilesPath).Workbook.Worksheets[2];
            for (int i = 0; i < amountOfValueToRead; i++)
            {
                valuesFromExcel.Enqueue(
                    excelWorksheet.GetValue(cellsPositionsXY[i, 1], cellsPositionsXY[i, 0])
                    .ToString()
                    );
            }
            //close excel file
            excelWorksheet.Dispose();

            if (valuesFromExcel.Count != amountOfValueToRead)
            {
                Console.WriteLine("Error: not enough values in excel file");
                CloseApp();
            }

            return valuesFromExcel.ToArray();
        }


        static void ChangingFiles(int carriersPerDay, int amountOfFiles, string directoryPath, string newDirectoryPath)
        {
            int[,] cellsPositionsXY = WitchCellsToChange();
            string [] value = null; ;
            //Variable without values
            string[] valuesFromExcel = ReadValuesFromExcel(
                FindFileWithExtension(directoryPath, "*.xlsx"));    //amount of value to change
            int amountOfValues = valuesFromExcel.Length;
            
            
            
            string filePath = FindFileWithExtension(directoryPath, "*.xlsx");   //bellow try for this variable
            string fileName = filePath.Substring(directoryPath.Length + 1);     //file name (cut path)
            string newFilePath = newDirectoryPath + $@"\{fileName}";             //coppy file in new directory

            // 1. Check file name
            //
            //Ważne sprawdzanie nazwy pliku pierwszego
            try
            {
                fileName.Substring(fileName.IndexOf("-") - 1, fileName.Length - fileName.IndexOf("-") + 1);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(
                    "Error" + "\n" +
                    "Check name of oryginal file. Should start like \"CAxxxx - ...\"" + "\n" +
                    "For example: \"CA0001 - RoDipE carrier precommissioning Check list v1.8\" "
                    );
                Console.ResetColor();

                //Open excel -> WorkSheet -> read one value and write to variable
                value[0] = new ExcelPackage(filePath).Workbook.Worksheets[2].GetValue(cellsPositionsXY[0, 0], cellsPositionsXY[0, 1]).ToString();
                fileName = $"CA{value[0]} - {fileName}";

                newFilePath = newDirectoryPath + $@"\{fileName}";
                Console.WriteLine("Changing name of first file on: " + fileName);
            }

            // 3.1. Info:
            Console.WriteLine($"\nCreated copy for: CA{value[0]}");

            // 2. Operate on excel file 
            File.Copy(filePath, newFilePath);                   //Copy Excel to new diercotry
            ExcelPackage excel = new ExcelPackage(newFilePath); //Create instance for excel
            ExcelWorksheet ws = excel.Workbook.Worksheets[2];   //Create instance for worksheet in excel

            int ipDevice_1;                                     // First part of IP for WLAN
            int ipDevice_2;                                      // First part of IP for PHC BOX
            string nameZeros;                               // First part of Name
            int nameNumber;                                 // Second part of Name

            //for (int i = 0; i < amountOfValues; i++)
            for (int i = 0; i < 4; i++)
            {
                value[i] = ws.GetValue(cellsPositionsXY[i, 0], cellsPositionsXY[i, 1]).ToString();   //GetValue() - takes value form cell

                //value[0] = 0001;          //Device_number
                //value[1] = 192;           //IP (last octet) - Device_1
                //value[2] = 102;           //IP (last octet) - Device_2
                //value[3] = dd.mm.yyyy;    //Date       
            }


            // 4. Prepare data
            // 4.1. Name for new file
            nameNumber = Int32.Parse(value[0]);
            nameZeros = value[0].Substring(0, (value[0].Length - 2));

            // 4.2. Conversion (casting) string to int
            ipDevice_1 = Int32.Parse(value[1]);
            ipDevice_2 = Int32.Parse(value[2]);

            // 4.3. Creating new variable "date" of type "DateTime". Parse string to DateTime
            DateTime.TryParse(value[3], out DateTime date);

            // 4.4. Edit file name
            Console.WriteLine(fileName);
            fileName = fileName.Substring(fileName.IndexOf("-") - 1, fileName.Length - fileName.IndexOf("-") + 1);

            if (carriersPerDay == 0)
            {
                date = date.AddDays(-1);
            }
            // 5. Changing data and create new excel
            for (int i = 1; i <= amountOfFiles; i++)
            {
                // 5.1. Change Name:
                if (nameNumber + i < 10) { value[0] = nameZeros + "0" + (nameNumber + i); }
                else { value[0] = nameZeros + (nameNumber + i); }

                // 5.2. Change and prepare IP:
                value[1] = (ipDevice_1 + i).ToString();   //WLAN
                value[2] = (ipDevice_2 + i).ToString();    //PHC

                // 5.3. Change date
                if ((i - 1) % carriersPerDay == 0)
                {
                    date = date.AddDays(1);
                }
                if (6 == (int)date.DayOfWeek)       //if Saturday
                {
                    date = date.AddDays(2);
                }
                else if (0 == (int)date.DayOfWeek)  //if Sunday 
                {
                    date = date.AddDays(1);
                }
                value[3] = date.ToString("d");      //"d" change to string but in data structure

                // 5.4. Write values in WorkSheet into correct cells (not saved yet)
                for (int j = 0; j < 4; j++)
                { ws.Cells[cellsPositionsXY[j, 0], cellsPositionsXY[j, 0]].Value = value[j]; }

                // 5.5. Save content to excel file  
                File.WriteAllBytes(newFilePath, excel.GetAsByteArray());

                // 5.6. Copy file and save with new name and changeds
                File.Copy(newFilePath, $@"{newDirectoryPath}\CA{value[0]}{fileName}");

                // 5.7. Info: new file created
                Console.WriteLine($"New excel for: CA{value[0]} is ready, with date: {date.ToShortDateString()}");
            };


            // 6. Clear all
            // 6.1. Close instance of excel (close file)
            excel.Dispose();                    //Clsoe excel instance

            // 6.2. Correcting the first copy
            File.Delete(newFilePath);           //Delete modified copy
            File.Copy(filePath, newFilePath);   //Create copy form orginal to new diercotry
        }
    }
}
