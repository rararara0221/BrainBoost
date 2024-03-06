using BrainBoost.Models;

namespace BrainBoost.Parameter
{
    public class QuestionList
    {
        public Question QuestionData { get; set; }
        public Option OptionData { get; set; }
        public List<string> Options { get; set; }
        public Answer AnswerData { get; set; }
    }
}
