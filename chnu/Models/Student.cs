﻿using System.Collections.Generic;

namespace chnu.Models
{
    public class Student
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public bool Debt { get; set; }

        public virtual List<Subject> Subjects { get; set; }
    }
}
