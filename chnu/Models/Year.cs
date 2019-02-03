using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace chnu.Models
{
    public class Year
    {
        public int Id { get; set; }

        public string Name { get; set; }
        // ex. 2014-2015, 2018-2019

        public bool IsSemestr { get; set; }

        public int PartOfYear { get; set; }

        public List<Group> Groups { get; set; }
    }
}
