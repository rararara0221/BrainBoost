namespace BrainBoost.Models
{
    public class Question
    {
        public int question_id { get; set; }
        public int type_id { get; set; }

        public float? question_grade { get; set; }

        public string question_content { get; set; }

        public string? question_picture {  get; set; }
    }
}
