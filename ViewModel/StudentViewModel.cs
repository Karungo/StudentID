using Microsoft.Win32;
using OfficeOpenXml;
using StudentID.Model;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using OpenCvSharp;
using System.Collections.Generic;

public class StudentViewModel : INotifyPropertyChanged
{
    private ObservableCollection<Student> _students;
    public ObservableCollection<Student> Students
    {
        get => _students;
        set
        {
            _students = value;
            OnPropertyChanged(nameof(Students));
        }
    }

    public ICommand LoadFileCommand { get; set; }
    public ICommand UploadPhotosCommand { get; set; }
    public ICommand ExportToFileCommand { get; set; }
    public ICommand CropPhotoCommand { get; set; }
    public ICommand DeletePhotoCommand { get; set; }
    public ICommand ReuploadPhotoCommand { get; set; }

    // Progress reporting and cancellation support
    private double _progress;
    public double Progress
    {
        get => _progress;
        set
        {
            _progress = value;
            OnPropertyChanged(nameof(Progress));
        }
    }

    private bool _isLoading;
    public bool IsLoading
    {
        get => _isLoading;
        set
        {
            _isLoading = value;
            OnPropertyChanged(nameof(IsLoading));
        }
    }

    private CancellationTokenSource _cancellationTokenSource;

    public StudentViewModel()
    {
        Students = new ObservableCollection<Student>();
        LoadFileCommand = new RelayCommand(async () => await LoadFileAsync());
        ExportToFileCommand = new RelayCommand(async () => await ExportToFileAsync());
        UploadPhotosCommand = new RelayCommand(async () => await UploadPhotosAsync());
        CropPhotoCommand = new RelayCommand<string>(async (admissionNumber) => await CropPhotoAsync(admissionNumber));
        DeletePhotoCommand = new RelayCommand<string>(DeletePhoto);
        ReuploadPhotoCommand = new RelayCommand<string>(async (admissionNumber) => await ReuploadPhotoAsync(admissionNumber));
        _cancellationTokenSource = new CancellationTokenSource();
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
            IsLoading = true;
            await Task.Run(() => ReadExcelFileInBatchesAsync(filePath));
            IsLoading = false;
        }
    }

    private async Task ReadExcelFileInBatchesAsync(string filePath)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        const int batchSize = 1000;
        FileInfo fileInfo = new FileInfo(filePath);
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

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

                if (regex.IsMatch(admissionNumber))
                {
                    string course = GetCourseFromAdmissionNumber(admissionNumber);
                    DateTime expiryDate = CalculateExpiryDate(admissionNumber);

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Students.Add(new Student
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
                    });
                }

                if (row % batchSize == 0)
                {
                    await Task.Delay(100); // Simulate batch processing
                }
            }
        }
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

    private async Task UploadPhotosAsync()
    {
        OpenFileDialog openFileDialog = new OpenFileDialog
        {
            DefaultExt = ".jpg",
            Filter = "Image Files (*.jpg, *.jpeg, *.png)|*.jpg;*.jpeg;*.png",
            Multiselect = true
        };

        bool? result = openFileDialog.ShowDialog();
        if (result == true)
        {
            IsLoading = true;
            await Task.WhenAll(openFileDialog.FileNames.Select(filePath => Task.Run(() =>
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath).ToUpper().Replace('-', '/');
                var student = Students.FirstOrDefault(s => s.AdmissionNumber == fileName);

                if (student != null)
                {
                    student.PhotoPath = ProcessPhoto(filePath);
                }
                else
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Students.Add(new Student
                        {
                            AdmissionNumber = fileName,
                            PhotoPath = ProcessPhoto(filePath)
                        });
                    });
                }
            })));

            IsLoading = false;
        }
    }

    private string ProcessPhoto(string filePath)
    {
        return ResizeAndCropPhoto(filePath);
    }

    private string ResizeAndCropPhoto(string filePath)
    {
        using (var srcImage = new Mat(filePath))
        {
            using (var grayImage = new Mat())
            {
                Cv2.CvtColor(srcImage, grayImage, ColorConversionCodes.BGR2GRAY);

                string haarCascadePath = "haarcascade_frontalface_default.xml";
                var faceCascade = new CascadeClassifier(haarCascadePath);

                OpenCvSharp.Rect[] faces = faceCascade.DetectMultiScale(
                    grayImage,
                    scaleFactor: 1.1,
                    minNeighbors: 5,
                    minSize: new OpenCvSharp.Size(100, 100));

                if (faces.Length > 0)
                {
                    OpenCvSharp.Rect faceRect = faces[0];

                    using (var faceImage = new Mat(srcImage, faceRect))
                    {
                        using (var resizedFace = new Mat())
                        {
                            Cv2.Resize(faceImage, resizedFace, new OpenCvSharp.Size(600, 600));

                            string processedPhotoPath = Path.Combine(
                                Path.GetDirectoryName(filePath),
                                "Processed_" + Path.GetFileName(filePath));

                            resizedFace.SaveImage(processedPhotoPath);
                            return processedPhotoPath;
                        }
                    }
                }
                else
                {
                    return filePath;
                }
            }
        }
    }

    private async Task CropPhotoAsync(string admissionNumber)
    {
        var student = Students.FirstOrDefault(s => s.AdmissionNumber == admissionNumber);
        if (student != null && !string.IsNullOrEmpty(student.PhotoPath))
        {
            student.PhotoPath = ResizeAndCropPhoto(student.PhotoPath);
            OnPropertyChanged(nameof(Students));
        }
    }

    private void DeletePhoto(string admissionNumber)
    {
        var student = Students.FirstOrDefault(s => s.AdmissionNumber == admissionNumber);
        if (student != null)
        {
            student.PhotoPath = null;
            OnPropertyChanged(nameof(Students));
        }
    }

    private async Task ReuploadPhotoAsync(string admissionNumber)
    {
        var student = Students.FirstOrDefault(s => s.AdmissionNumber == admissionNumber);
        if (student != null)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                DefaultExt = ".jpg",
                Filter = "Image Files (*.jpg, *.jpeg, *.png)|*.jpg;*.jpeg;*.png",
                Multiselect = false
            };

            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                student.PhotoPath = ProcessPhoto(openFileDialog.FileName);
                OnPropertyChanged(nameof(Students));
            }
        }
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
            IsLoading = true;
            await Task.Run(() => ExportDataToExcel(filePath, Students.ToList()));
            IsLoading = false;
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
            worksheet.Cells[1, 9].Value = "Photo";

            int row = 2;
            foreach (var student in students)
            {
                worksheet.Cells[row, 1].Value = student.Id;
                worksheet.Cells[row, 2].Value = student.Name;
                worksheet.Cells[row, 3].Value = student.Gender;
                worksheet.Cells[row, 4].Value = student.AdmissionNumber;
                worksheet.Cells[row, 5].Value = student.IdNumber;
                worksheet.Cells[row, 6].Value = student.Course;
                worksheet.Cells[row, 7].Value = student.Nationality;
                worksheet.Cells[row, 8].Value = student.ExpiryDate.ToString("yyyy");
                if (!string.IsNullOrEmpty(student.PhotoPath) && File.Exists(student.PhotoPath))
                {
                    var imageFileInfo = new FileInfo(student.PhotoPath);
                    var excelImage = worksheet.Drawings.AddPicture($"Photo_{row}", imageFileInfo);
                    excelImage.SetPosition(row - 1, 0, 8, 0);
                    excelImage.SetSize(50, 50);
                }

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
