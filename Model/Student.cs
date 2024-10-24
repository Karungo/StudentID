using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace StudentID.Model
{
    public class Student
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Gender { get; set; }
        public string Nationality { get; set; }
        public string AdmissionNumber { get; set; }
        public string IdNumber { get; set; }
        public string Course { get; set; }
        public DateTime ExpiryDate { get; set; }
        public string PhotoPath{ get; set; }
        public ImageSource PhotoImage{ get; set; }
        public string SelectedAdmissionNumber { get; set; }
    }
}

