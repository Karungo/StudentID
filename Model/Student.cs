﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
