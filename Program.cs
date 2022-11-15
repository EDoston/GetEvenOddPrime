using OfficeOpenXml;
using System.Data;

namespace GetEvenOddAndPrimeNumbers
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string txt = File.ReadAllText(@"list_of_random_numbers.txt");

                var randomNumbers = txt.Split(",").Select(x => int.Parse(x));

                var numbers = new Numbers();

                numbers.even = randomNumbers.Where(x => x % 2 == 0);
                numbers.odd = randomNumbers.Where(x => x % 2 != 0);
                numbers.prime = GetPrimeNumbers(randomNumbers);

                numbers.odd = numbers.odd.Except(numbers.prime);

                SaveInExcel(numbers);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        static IEnumerable<int> GetPrimeNumbers(IEnumerable<int> numbers) 
        {
            var primeNumbers = new List<int>();

            foreach (var item in numbers)
                if (IsNumberPrime(item))
                    primeNumbers.Add(item);

            return primeNumbers;
        }

        static bool IsNumberPrime(int n)
        {
            if (n == 1) return false;

            for (int i = 2; i <= Math.Sqrt(n); i++)
                if (n % i == 0)
                    return false;

            return true;
        }

        static void SaveInExcel(Numbers evenOddPrime)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var file = new FileInfo(@"Result.xlsx");

                if (file.Exists)
                    file.Delete();

                using var package = new ExcelPackage(file);

                var ws = package.Workbook.Worksheets.Add("Even, odd and prime numbers");

                var firstColumn = ws.Cells["A1"];
                firstColumn.Value = "Even numbers";
                firstColumn.AutoFitColumns();

                ws.Cells["A2"].LoadFromCollection(evenOddPrime.even);

                var secondColumn = ws.Cells["B1"];
                secondColumn.Value = "Odd numbers";
                secondColumn.AutoFitColumns();

                ws.Cells["B2"].LoadFromCollection(evenOddPrime.odd);

                var thirdColumn = ws.Cells["C1"];
                thirdColumn.Value = "Prime numbers";
                thirdColumn.AutoFitColumns();

                ws.Cells["C2"].LoadFromCollection(evenOddPrime.prime);

                package.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}