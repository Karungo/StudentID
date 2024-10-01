using Microsoft.Win32;
using OfficeOpenXml;
using StudentID.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

public class StudentViewModel : INotifyPropertyChanged
{
    private List<Student> _students;
    public List<Student> Students
    {
        get => _students;
        set
        {
            _students = value;
            OnPropertyChanged(nameof(Students));
        }
    }

    public ICommand LoadFileCommand { get; set; }
    public ICommand ExportToFileCommand { get; set; }

    public StudentViewModel()
    {
        LoadFileCommand = new RelayCommand(async () => await LoadFileAsync());
        ExportToFileCommand = new RelayCommand(async () => await ExportToFileAsync());
    }

    private async Task LoadFileAsync()
    {
        OpenFileDialog openFileDialog = new OpenFileDialog
        {
            DefaultExt = ".xlsx",
            Filter = "Excel Files (*.xlsx)|*.xlsx"
        };

        bool? result = openFileDialog.ShowDialog();
        if (result == true)
        {
            string filePath = openFileDialog.FileName;
            await Task.Run(() => Students = ReadExcelFileInBatches(filePath));
        }
    }

    private List<Student> ReadExcelFileInBatches(string filePath)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        List<Student> students = new List<Student>();
        const int batchSize = 1000; // Process in batches for scalability

        FileInfo fileInfo = new FileInfo(filePath);
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            // Define the pattern for AdmissionNumber
            string pattern = @"^[A-Za-z]{3}/\d{3}/[A-Za-z]\d{2}/\d{3}$";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

            for (int row = 2; row <= rowCount; row++)
            {
                string no = worksheet.Cells[row, 1].Text;
                string name = worksheet.Cells[row, 2].Text;
                string gender = worksheet.Cells[row, 3].Text;
                string admissionNumber = worksheet.Cells[row, 4].Text;
                string idNumber = worksheet.Cells[row, 5].Text;
                string nationality = worksheet.Cells[row, 6].Text;

                // Check if the admission number matches the regex pattern
                if (regex.IsMatch(admissionNumber))
                {
                    string course = GetCourseFromAdmissionNumber(admissionNumber);
                    DateTime expiryDate = CalculateExpiryDate(admissionNumber);

                    students.Add(new Student
                    {
                        Id = no,
                        Name = name.ToUpper(),
                        Gender = gender.ToUpper(),
                        AdmissionNumber = admissionNumber.ToUpper(),
                        IdNumber = idNumber.ToUpper(),
                        Course = course.ToUpper(),
                        Nationality = nationality.ToUpper(),
                        ExpiryDate = expiryDate
                    });
                }
                else
                {
                    // If admission number doesn't match, skip this entry
                    continue;
                }

                if (row % batchSize == 0)
                {
                    // Optionally clear students after each batch to manage memory
                    students.Clear();
                }
            }
        }
        return students;
    }

    private string GetCourseFromAdmissionNumber(string admissionNumber)
    {
        return admissionNumber.Split('/')[0];
    }

    private DateTime CalculateExpiryDate(string admissionNumber)
    {
        string[] parts = admissionNumber.Split('/');
        int duration = int.Parse(parts[1]);

        DateTime expiryDate = DateTime.Now;
        switch (duration)
        {
            case 600:
                expiryDate = expiryDate.AddYears(3);
                break;
            case 500:
                expiryDate = expiryDate.AddYears(2);
                break;
            case 400:
                expiryDate = expiryDate.AddYears(1);
                break;
            case 300:
                expiryDate = expiryDate.AddMonths(3);
                break;
        }
        return expiryDate;
    }

    private async Task ExportToFileAsync()
    {
        SaveFileDialog saveFileDialog = new SaveFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx",
            DefaultExt = ".xlsx"
        };

        bool? result = saveFileDialog.ShowDialog();
        if (result == true)
        {
            string filePath = saveFileDialog.FileName;
            await Task.Run(() => ExportDataToExcel(filePath, Students));
            MessageBox.Show("Data exported successfully!");
        }
    }

    private void ExportDataToExcel(string filePath, List<Student> students)
    {
        FileInfo file = new FileInfo(filePath);
        using (ExcelPackage package = new ExcelPackage(file))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Students");
            worksheet.Cells[1, 1].Value = "Id";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "Gender";
            worksheet.Cells[1, 4].Value = "Admission Number";
            worksheet.Cells[1, 5].Value = "ID Number";
            worksheet.Cells[1, 6].Value = "Course";
            worksheet.Cells[1, 7].Value = "Nationality";
            worksheet.Cells[1, 8].Value = "Expiry Date";

            int row = 2;
            foreach (var student in students)
            {
                worksheet.Cells[row, 1].Value = student.Id; 
                worksheet.Cells[row, 2].Value = student.Name.ToUpper();
                worksheet.Cells[row, 3].Value = student.Gender.ToUpper();
                worksheet.Cells[row, 4].Value = student.AdmissionNumber.ToUpper();
                worksheet.Cells[row, 5].Value = student.IdNumber.ToUpper();
                worksheet.Cells[row, 6].Value = student.Course.ToUpper();
                worksheet.Cells[row, 7].Value = student.Nationality.ToUpper();
                worksheet.Cells[row, 8].Value = student.ExpiryDate.ToString("yyyy");
                row++;
            }

            package.Save();
        }
    }

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
