using OfficeOpenXml;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace ExcelToNewExcelsCreator
{

    internal class Program
    {
        public static class CellsToReadChange
        {
            private static int[,] cellPostitionXY = {
                { 6, 3  },  //Device_number
                { 7, 11 },  //IP (last octet) - Device_1
                { 7, 13 },  //IP (last octet) - Device_2
                { 6, 23 }   //Date
            };

            public static int[,] GetCellsPostition => cellPostitionXY;
        }
        static void Main(string[] args)
        {
            //Package licence
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //Paths
            string baseFilesDirectoryPath = GetPathToDirectory("Files");
            string newFilesDirectoryPath = baseFilesDirectoryPath + @"\NewFiles";
            baseFilesDirectoryPath += @"\BaseFiles";

            //Check if files exist, name correct, can read value form excel
            string baseExcelFile = FindFileWithExtension(baseFilesDirectoryPath, "*.xlsx");
            string[] fileNameIn3Parts = ReturnFileNameIn3PartsIfCorrect(baseExcelFile);
            string[] valuesR​eadeFromExcel = ReadValuesFromExcel(baseExcelFile);

            CreateNewDirectory(newFilesDirectoryPath);
            try
            {
                File.Copy(baseExcelFile, @$"{newFilesDirectoryPath}\{fileNameIn3Parts[0]}{fileNameIn3Parts[1]}{fileNameIn3Parts[2]}");
            }
            catch 
            {
                Console.WriteLine("Can't copy file to new directory");
                CloseApp(); 
            }


            // Ask User
            Console.WriteLine("Write below how many NEW files create");
            int amountOfNewFiles = AskUserAboutAmount();
            if (amountOfNewFiles == 0)
            {
                Console.WriteLine("You chose 0 new files");
                CloseApp();
            }
            Console.WriteLine();

            Console.WriteLine("Write below how many carriers per day");
            int carriersPerDay = AskUserAboutAmount();
            if (carriersPerDay == 0)
            {
                Console.WriteLine("You chose zero, all files with same date");
            }
            Console.WriteLine();

            string[] newExcelFilesPath = PreaperExcelFileName(fileNameIn3Parts, newFilesDirectoryPath, amountOfNewFiles);

            CreateNewExcelFiles(baseExcelFile, newExcelFilesPath);
            ChangeValuesInNewExcelFiles(newExcelFilesPath, valuesReadeFromExcel, carriersPerDay);


            Console.WriteLine($"\nAll new files saved in directory:" +
                $"\n{newFilesDirectoryPath}"+
                "Press any key to close the window");
            Console.ReadKey();

            return;
        }


        enum ColorMessageEnum
        {
            yellow,
            red,
        }
        static void ColoredMesage(ColorMessageEnum color, string message)
        {
            switch (color)
            {
                case ColorMessageEnum.yellow:

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;

                case ColorMessageEnum.red:

                    Console.ForegroundColor = ConsoleColor.Red;
                    break;

                default:

                    Console.ResetColor();
                    break;
            }
            Console.WriteLine(message);
            Console.ResetColor();
            return;
        }
        static void CloseApp([CallerMemberName] string methodName = "")
        {
            Debug.WriteLine($"!!!There was a problem in method: {methodName}");

            Console.WriteLine("\nPress Enter to close app");
            Console.ReadKey();
            Environment.Exit(0);
        }


        static string GetPathToDirectory()
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
            }
            return path;
        }
        static string GetPathToDirectory(string directoryName)
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
                path = path + @"\" + directoryName;
            }
            return path;
        }
        static string FindFileWithExtension(string directoryPath, string extension)
        {
            string[] path = null;
            try
            {
                path = Directory.GetFiles(directoryPath, extension);
            }
            catch (DirectoryNotFoundException)
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
                    if ('~' != path[i][0])
                    {
                        return path[i];
                    }
                }
            }

            Console.WriteLine("Valid file not found, check directory:\n" + directoryPath);
            CloseApp();
            return null;
        }
        static void CreateNewDirectory(string newDirectoryPath)
        {
            try
            {
                if (Directory.Exists(newDirectoryPath))
                {
                    Directory.Delete(newDirectoryPath, true);
                    Directory.CreateDirectory(newDirectoryPath);
                }
                else
                {
                    Directory.CreateDirectory(newDirectoryPath);
                }
            }
            catch (UnauthorizedAccessException)
            {
                Debug.WriteLine("Catch in CreateNewDirectory -> UnauthorizedAccessException");

                Console.WriteLine("No permition to delete directory:\n" + newDirectoryPath);
                CloseApp();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Can't delete or create new directory");
                Console.WriteLine($"Close all files from directory:\n" + newDirectoryPath);
                Console.ResetColor();

                CloseApp();
            }
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
                ColoredMesage(ColorMessageEnum.yellow, "Error: not valid integer value");
                CloseApp();
            }
            catch
            {
                ColoredMesage(ColorMessageEnum.red, "!!!Error durring reading value from keyboard");
                CloseApp();
            }

            return 0;
        }


        static string[] ReturnFileNameIn3PartsIfCorrect(string baseExcelFile)
        {
            if (baseExcelFile == "")
            {
                Console.WriteLine("Excel file not found");
                CloseApp();
            }

            string fullFileName = "";
            try
            {
                fullFileName =
                baseExcelFile.Substring(
                baseExcelFile.LastIndexOf(@"\") + 1);
            }
            catch
            {
                Debug.WriteLine("Error in CheckFileName -> try");
                CloseApp();
            }

            int amountOfLetters = 0;
            if (Char.IsLetter(fullFileName[amountOfLetters]))
            {
                amountOfLetters++;
            }
            else
            {
                Debug.WriteLine("Error in CheckFileName -> First char isn't letter");
                Console.WriteLine("First char isn't letter");
                Internal_CloseWithComment();
            }

            while (Char.IsLetter(fullFileName[amountOfLetters]))
            {
                amountOfLetters++;
            }

            string fileLetters = "";
            string fileNumber = ""; ///// usunąć spacje
            string fileDescription = "";
            try
            {
                fileLetters = fullFileName.Substring(0, amountOfLetters);
                fileNumber = fullFileName
                    .Substring(amountOfLetters, (fullFileName.IndexOf("-")-2))
                    .Trim();
                fileDescription = " " + fullFileName.Substring(fullFileName.IndexOf("-"));
            }
            catch
            {
                Debug.WriteLine("Problem with substrings");
                Internal_CloseWithComment();
            }

            try
            {
                if (0 > Int32.Parse(fileNumber))
                {
                    Internal_CloseWithComment();
                }
            }
            catch
            {
                Debug.WriteLine("Error in CheckFileName -> Int32.Parse(fileNumber)");
                Internal_CloseWithComment();
            }

            return new string[] { fileLetters, fileNumber, fileDescription };

            static void Internal_CloseWithComment()
            {
                Console.WriteLine("Check name of base excel file. It should start as follow 'CAxxx'");
                CloseApp();
            }
        }
        static string[] PreaperExcelFileName(string[] fileNameIn3Parts, string newFilesDirectoryPath, int amountOfNewFiles)
        {
            int localNumber = 0;
            try
            {
                localNumber = Int32.Parse(fileNameIn3Parts[1]);
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

            int orginalNameNumberLenght = fileNameIn3Parts[1].Length;
            string[] newExcelFilesPath = new string[amountOfNewFiles];
            for (int i = 0; i < amountOfNewFiles; i++)
            {
                newExcelFilesPath[i] = (localNumber + (i + 1)).ToString();
                newExcelFilesPath[i] =
                    newFilesDirectoryPath + @"\" +
                    fileNameIn3Parts[0] +
                    new string('0', orginalNameNumberLenght - newExcelFilesPath[i].Length) +
                    newExcelFilesPath[i] +
                    fileNameIn3Parts[2];
            }

            return newExcelFilesPath;
        }
        static void CreateNewExcelFiles(string baseExcelFile, string[] newExelFilesPath)
        {
            int amountOfNewFiles = newExelFilesPath.Length;
            try
            {
                for (int i = 0; i < amountOfNewFiles; i++)
                {
                    File.Copy(baseExcelFile, newExelFilesPath[i]);
                }
            }
            catch
            {
                CloseApp();
            }
            return;
        }

        static string[] ReadValuesFromExcel(string excelFilesPath)
        {
            Queue<string> valuesFromExcel = new Queue<string>();
            int[,] cellsPositionXY = CellsToReadChange.GetCellsPostition;
            int amountOfValueToRead = cellsPositionXY.Length / 2;

            try
            {
                ExcelWorksheet excelWorksheet = new ExcelPackage(excelFilesPath).Workbook.Worksheets[2];

                for (int i = 0; i < amountOfValueToRead; i++)
                {
                    valuesFromExcel.Enqueue(
                        excelWorksheet.GetValue(cellsPositionXY[i, 1], cellsPositionXY[i, 0])
                        .ToString()
                        );
                }

                excelWorksheet.Dispose();
            }
            catch
            {
                Debug.WriteLine("Can't open worsheet nr.2");
                CloseApp();
            }


            if (valuesFromExcel.Count != amountOfValueToRead)
            {
                Console.WriteLine("Error: not enough values in excel file");
                CloseApp();
            }

            return valuesFromExcel.ToArray();
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
            for (int i = 0; i < amountOfNewFiles; i++)
            {
                newIpAddress[i] = (localIpAddress + (i + 1)).ToString();
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
            for (int i = 0; i < amountOfNewFiles; i++)
            {
                if ((i + 1) % carriersPerDay == 0)
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

        static void ChangeValuesInNewExcelFiles(string[] newExcelFilesPath, string[] valuesReadeFromExcel, int carriersPerDay)
        {
            int amountOfNewFiles = newExcelFilesPath.Length;
            string[] valueFileNumber = ChangeNumber(valuesReadeFromExcel[0], amountOfNewFiles);
            string[] valueIpAddress_1 = ChangeIpAddress(valuesReadeFromExcel[1], amountOfNewFiles);
            string[] valueIpAddress_2 = ChangeIpAddress(valuesReadeFromExcel[2], amountOfNewFiles);
            string[] valueDate = ChangeDate(valuesReadeFromExcel[3], amountOfNewFiles, carriersPerDay);

            try
            {
                int[,] cellsPositionXY = CellsToReadChange.GetCellsPostition;
                for (int i = 0; i < amountOfNewFiles; i++)
                {
                    ExcelPackage localExcel = new ExcelPackage(newExcelFilesPath[i]);
                    ExcelWorksheet localWorksheets = localExcel.Workbook.Worksheets[2];

                    localWorksheets.Cells[cellsPositionXY[0, 1], cellsPositionXY[0, 0]].Value = valueFileNumber[i];
                    localWorksheets.Cells[cellsPositionXY[1, 1], cellsPositionXY[1, 0]].Value = valueIpAddress_1[i];
                    localWorksheets.Cells[cellsPositionXY[2, 1], cellsPositionXY[2, 0]].Value = valueIpAddress_2[i];
                    localWorksheets.Cells[cellsPositionXY[3, 1], cellsPositionXY[3, 0]].Value = valueDate[i];


                    File.WriteAllBytes(newExcelFilesPath[i], localExcel.GetAsByteArray());
                    localExcel.Dispose();
                    Console.WriteLine($"{i + 1} new file created");
                }
            }
            catch
            {
                CloseApp();
            }

        }

    }
}
