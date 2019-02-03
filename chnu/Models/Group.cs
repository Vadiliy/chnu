using System.Collections.Generic;

namespace chnu.Models
{
    public class Group
    {
        public int Id { get; set; }

        public string NameGroup { get; set; }

        public virtual List<Student> Students { get; set; }
    }
}
