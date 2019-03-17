namespace chnu.Models
{
    public class Subject
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string ExamType { get; set; }

        public double Credits { get; set; }

        public double Hours { get; set; }

        public string Lecture { get; set; }

        public int Score { get; set; }

        public string DateDebt { get; set; }

        public virtual Student Student { get; set; }
    }
}
