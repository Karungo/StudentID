using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace StudentID.Model
{
    public class PhotoModel : INotifyPropertyChanged
    {
        private string _photoPath;
        private BitmapImage _photoImage;

        public string PhotoPath
        {
            get => _photoPath;
            set
            {
                _photoPath = value;
                PhotoImage = LoadImage(_photoPath); // Load the image without locking the file
                OnPropertyChanged(nameof(PhotoPath));
            }
        }

        public BitmapImage PhotoImage
        {
            get => _photoImage;
            set
            {
                _photoImage = value;
                OnPropertyChanged(nameof(PhotoImage));
            }
        }
        public BitmapImage LoadImage(string path)
        {
            var bitmap = new BitmapImage();
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad; // Load the image into memory
                bitmap.StreamSource = stream;
                bitmap.EndInit();
            }
            bitmap.Freeze(); // Makes the image usable across threads
            return bitmap;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
