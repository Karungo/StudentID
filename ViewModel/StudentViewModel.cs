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
using System.Windows.Controls;
using System.Diagnostics;
using System.Windows.Media.Imaging;
using System.Windows.Media;

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
    private ObservableCollection<string> _admissionNumbers;
    public ObservableCollection<string> AdmissionNumbers
    {
        get { return _admissionNumbers; }
        set
        {
            _admissionNumbers = value;
            OnPropertyChanged(nameof(AdmissionNumbers));
        }
    }

    private string _selectedAdmissionNumber;
    public string SelectedAdmissionNumber
    {
        get { return _selectedAdmissionNumber; }
        set
        {
            if (_selectedAdmissionNumber != value)
            {
                _selectedAdmissionNumber = value;
                OnPropertyChanged(nameof(SelectedAdmissionNumber));
            }
        }
    }

    public ICommand LoadFileCommand { get; set; }
    public ICommand UploadPhotosCommand { get; set; }
    public ICommand LoadPhotosCommand { get; set; }
    public ICommand ExportToFileCommand { get; set; }
    public ICommand CropPhotoCommand { get; set; }
    public ICommand DeletePhotoCommand { get; set; }
    public ICommand ReuploadPhotoCommand { get; set; }
    public ICommand AddAdmissionNumbersCommand { get; }
    public ICommand RenamePhotoCommand { get; private set; }

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
        AdmissionNumbers = new ObservableCollection<string>();
        AddAdmissionNumbersCommand = new RelayCommand(() => AddAdmissionNumbers());
        LoadFileCommand = new RelayCommand(async () => await LoadFileAsync());
        ExportToFileCommand = new RelayCommand(async () => await ExportToFileAsync());
        UploadPhotosCommand = new RelayCommand(async () => await UploadPhotosAsync());
        LoadPhotosCommand = new RelayCommand(async () => await LoadPhotosAsync());
        CropPhotoCommand = new RelayCommand<string>(async (admissionNumber) => await CropPhotoAsync(admissionNumber));
        DeletePhotoCommand = new RelayCommand<string>(DeletePhoto);
        ReuploadPhotoCommand = new RelayCommand<string>(async (admissionNumber) => await ReuploadPhotoAsync(admissionNumber));
        RenamePhotoCommand = new RelayCommand<Student>(RenamePhoto);
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
                string gender = worksheet.Cells[row, 3].Text.ToUpper();
                string admissionNumber = worksheet.Cells[row, 4].Text;
                string idNumber = worksheet.Cells[row, 5].Text;
                string nationality = worksheet.Cells[row, 6].Text;

                if (gender == "M")
                {
                    gender = "MALE";
                }
                else if (gender == "F")
                {
                    gender = "FEMALE";
                }
                else if (gender != "MALE" && gender != "FEMALE")
                {
                    gender = "Undefined";
                }

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
                            Gender = gender,
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

            // Limit concurrency with a Task list
            var uploadTasks = openFileDialog.FileNames.Select(filePath => Task.Run(() =>
            {
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(filePath).ToUpper().Replace('-', '/');
                    var student = Students.FirstOrDefault(s => s.AdmissionNumber == fileName);

                    string optimizedPhotoPath = ProcessPhoto(filePath); // Ensure memory-friendly processing
                    var photoImage = LoadImageFromPath(optimizedPhotoPath);

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (student != null)
                        {
                            student.PhotoPath = optimizedPhotoPath;
                            student.PhotoImage = photoImage;
                        }
                        else
                        {
                            Students.Add(new Student
                            {
                                AdmissionNumber = fileName,
                                PhotoPath = optimizedPhotoPath,
                                PhotoImage = photoImage
                            });
                        }
                    });
                }
                catch (Exception ex)
                {
                    // Log or handle exceptions as needed
                    Console.WriteLine($"Error uploading photo for {filePath}: {ex.Message}");
                }
            }));

            // Await tasks with controlled concurrency
            await Task.WhenAll(uploadTasks);

            IsLoading = false;
        }
    }
    private async Task LoadPhotosAsync()
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

            // Limit concurrency with a Task list
            var uploadTasks = openFileDialog.FileNames.Select(filePath => Task.Run(() =>
            {
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(filePath).ToUpper().Replace('-', '/');
                    var student = Students.FirstOrDefault(s => s.AdmissionNumber == fileName);

                    var photoImage = LoadImageFromPath(filePath);

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (student != null)
                        {
                            student.PhotoPath = filePath;
                            student.PhotoImage = photoImage;
                        }
                        else
                        {
                            Students.Add(new Student
                            {
                                AdmissionNumber = fileName,
                                PhotoPath = filePath,
                                PhotoImage = photoImage
                            });
                        }
                    });
                }
                catch (Exception ex)
                {
                    // Log or handle exceptions as needed
                    Console.WriteLine($"Error uploading photo for {filePath}: {ex.Message}");
                }
            }));

            // Await tasks with controlled concurrency
            await Task.WhenAll(uploadTasks);

            IsLoading = false;
        }
    }
    private string ProcessPhoto(string filePath)
    {
        return ResizeAndCropPhoto(filePath);
    }
    private static bool _errorDisplayed = false;
    private string ResizeAndCropPhoto(string filePath)
    {
        // Construct the path for the haarcascade file
        string programFilesDir = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        string haarCascadePath = Path.Combine(programFilesDir, "Karungo", "StudentIDInstaller", "haarcascade_frontalface_default.xml");

        if (!File.Exists(haarCascadePath))
        {
            // Check if the error dialog has already been shown
            if (!_errorDisplayed)
            {
                _errorDisplayed = true;
                Application.Current.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"Haarcascade file not found at {haarCascadePath}. The application will now close.",
                                    "File Not Found",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                });

                Application.Current.Shutdown();
            }
            return "Not found";
        }

        using (var srcImage = new Mat(filePath))
        {
            using (var grayImage = new Mat())
            {
                Cv2.CvtColor(srcImage, grayImage, ColorConversionCodes.BGR2GRAY);

                // Load the cascade classifier from the file
                var faceCascade = new CascadeClassifier(haarCascadePath);

                OpenCvSharp.Rect[] faces = faceCascade.DetectMultiScale(
                    grayImage,
                    scaleFactor: 1.1,
                    minNeighbors: 5,
                    minSize: new OpenCvSharp.Size(100, 100));

                if (faces.Length > 0)
                {
                    OpenCvSharp.Rect faceRect = faces[0];

                    // Add padding to the face rectangle for passport-style photo
                    int verticalPadding = (int)(faceRect.Height * 0.4);
                    int horizontalPadding = (int)(faceRect.Width * 0.25);

                    // Adjust the rectangle to include padding
                    int newX = Math.Max(0, faceRect.X - horizontalPadding);
                    int newWidth = Math.Min(srcImage.Cols - newX, faceRect.Width + 2 * horizontalPadding);

                    int newY = Math.Max(0, faceRect.Y - verticalPadding);
                    int newHeight = Math.Min(srcImage.Rows - newY, faceRect.Height + 2 * verticalPadding);

                    OpenCvSharp.Rect adjustedRect = new OpenCvSharp.Rect(newX, newY, newWidth, newHeight);

                    // Crop the image based on the adjusted rectangle
                    using (var croppedImage = new Mat(srcImage, adjustedRect))
                    {
                        int targetWidthPixels = 236;
                        int targetHeightPixels = 300;

                        using (var resizedFace = new Mat())
                        {
                            Cv2.Resize(croppedImage, resizedFace, new OpenCvSharp.Size(targetWidthPixels, targetHeightPixels));

                            using (var finalImage = new Mat(new OpenCvSharp.Size(targetWidthPixels, targetHeightPixels), MatType.CV_8UC3, Scalar.White))
                            {
                                int xOffset = (targetWidthPixels - resizedFace.Width) / 2;
                                int yOffset = (targetHeightPixels - resizedFace.Height) / 2;

                                resizedFace.CopyTo(finalImage[new OpenCvSharp.Rect(xOffset, yOffset, resizedFace.Width, resizedFace.Height)]);

                                string processedPhotoPath = Path.Combine(
                                    Path.GetDirectoryName(filePath),
                                    "Processed_" + Path.GetFileName(filePath));

                                finalImage.SaveImage(processedPhotoPath);
                                return processedPhotoPath;
                            }
                        }
                    }
                }
                else
                {
                    // Log if no face is detected
                    Console.WriteLine($"No face detected in {filePath}. Returning original photo.");
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
            OnPropertyChanged(nameof(Students)); // Notify UI that the Students collection has changed
            MessageBox.Show("Photo has been deleted successfully.", "Deletion Success", MessageBoxButton.OK, MessageBoxImage.Information);
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
                student.PhotoPath = ProcessPhoto(openFileDialog.FileName); // Process and set new photo path
                OnPropertyChanged(nameof(Students)); // Notify UI that the Students collection has changed

                MessageBox.Show("Photo has been re-uploaded successfully.", "Upload Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
    private void AddAdmissionNumbers()
    {
        // Prompt user for input using a dialog box (or input box logic)
        var input = ShowInputDialog("Enter Admission Numbers (one per line):");

        if (!string.IsNullOrWhiteSpace(input))
        {
            // Split input by new lines and add to the collection
            var numbers = input.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var number in numbers)
            {
                // Add only if it's not already in the collection
                if (!AdmissionNumbers.Contains(number.Trim()))
                {
                    AdmissionNumbers.Add(number.Trim());
                }
            }
            AdmissionNumbers = new ObservableCollection<string>(AdmissionNumbers.OrderBy(x => x));
        }
    }
    // Function to display an input dialog and return the user's input
    public static string ShowInputDialog(string prompt)
    {
        // Create a window for user input
        System.Windows.Window inputWindow = new System.Windows.Window
        {
            Title = prompt,
            Width = 400,
            Height = 300,
            WindowStartupLocation = WindowStartupLocation.CenterScreen
        };

        StackPanel panel = new StackPanel();

        // TextBox wrapped in ScrollViewer to allow scrolling for large input
        ScrollViewer scrollViewer = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };

        TextBox inputBox = new TextBox
        {
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            MaxHeight = 200, // Set a max height to trigger scrolling when needed
            MinHeight = 30, // Set a minimum height for better user experience
            Height = Double.NaN // Auto adjust to fit the content until max height is reached
        };

        scrollViewer.Content = inputBox;
        panel.Children.Add(scrollViewer);

        Button okButton = new Button
        {
            Content = "OK",
            Width = 50,
            Margin = new Thickness(5),
            Background = (SolidColorBrush)new BrushConverter().ConvertFromString("#007ACC"),   // Set blue background
            Foreground = new SolidColorBrush(Colors.White),  // Set white text color
            BorderBrush = (SolidColorBrush)new BrushConverter().ConvertFromString("#007ACC"),  // Blue border (same as background for seamless look)
            BorderThickness = new Thickness(1),
            Padding = new Thickness(10),
            FontWeight = FontWeights.Bold,                   // Make text bold for visibility
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        // Define the ControlTemplate with rounded corners
        var template = new ControlTemplate(typeof(Button));
        var borderFactory = new FrameworkElementFactory(typeof(Border));
        borderFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
        borderFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
        borderFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
        borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(10)); // Rounded corners

        var contentPresenterFactory = new FrameworkElementFactory(typeof(ContentPresenter));
        contentPresenterFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        contentPresenterFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
        contentPresenterFactory.SetValue(ContentPresenter.ContentProperty, new TemplateBindingExtension(Button.ContentProperty));

        borderFactory.AppendChild(contentPresenterFactory);
        template.VisualTree = borderFactory;

        okButton.Template = template;
        okButton.Click += (sender, e) => { inputWindow.DialogResult = true; inputWindow.Close(); };

        panel.Children.Add(okButton);
        inputWindow.Content = panel;

        inputWindow.ShowDialog();
        return inputBox.Text;
    }
    private void RenamePhoto(Student student)
    {
        if (student == null || string.IsNullOrEmpty(student.PhotoPath))
        {
            MessageBox.Show("Invalid student or photo path.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (!string.IsNullOrEmpty(student.SelectedAdmissionNumber))
        {
            string newFileName = student.SelectedAdmissionNumber + Path.GetExtension(student.PhotoPath);
            string newFilePath = Path.Combine(Path.GetDirectoryName(student.PhotoPath), newFileName);

            try
            {
                if (File.Exists(student.PhotoPath))
                {
                    // Move the photo and rename it
                    File.Move(student.PhotoPath, newFilePath);
                    student.PhotoPath = newFilePath;

                    // Update PhotoImage for UI binding
                    student.PhotoImage = LoadImageFromPath(newFilePath);
                    OnPropertyChanged(nameof(student.PhotoImage));

                    // Update admission number and notify UI of the change
                    student.AdmissionNumber = student.SelectedAdmissionNumber;
                    AdmissionNumbers.Remove(student.SelectedAdmissionNumber);
                    OnPropertyChanged(nameof(Students));

                    // Notify user of success
                    MessageBox.Show($"Photo has been renamed and admission number updated to {student.SelectedAdmissionNumber}.",
                                    "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("The original photo file could not be found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error renaming file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
    private BitmapImage LoadImageFromPath(string imagePath)
    {
        BitmapImage bitmap = new BitmapImage();
        bitmap.BeginInit();
        bitmap.CacheOption = BitmapCacheOption.OnLoad; // Ensure image is fully loaded to avoid file locking issues
        bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
        bitmap.EndInit();
        bitmap.Freeze(); // Freeze for thread safety
        return bitmap;
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
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

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
                    // Replace '/' with '-' in the admission number
                    string sanitizedAdmissionNumber = student.AdmissionNumber.Replace("/", "-");

                    // Create a processed photo path using the sanitized admission number
                    string processedPhotoPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
                        "Processed_" + sanitizedAdmissionNumber + ".jpeg"
                    );

                    // Optionally, copy or process the image if necessary
                    File.Copy(student.PhotoPath, processedPhotoPath, true);

                    // Save the processed photo path in the Excel cell
                    worksheet.Cells[row, 9].Value = processedPhotoPath;
                }
                else
                {
                    worksheet.Cells[row, 9].Value = "";
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
