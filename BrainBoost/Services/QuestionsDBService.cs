using BrainBoost.Models;
using BrainBoost.Parameter;
using System.Data.SqlClient;
using Dapper;
using System.Text;
using System.Drawing.Printing;

namespace BrainBoost.Services
{
    public class QuestionsDBService
    {
        // 宣告Service
        private readonly string? cnstr;

        public QuestionsDBService(IConfiguration configuration)
        {
            cnstr = configuration.GetConnectionString("ConnectionName");
        }


        // 儲存題目
        public void InsertQuestion(QuestionList questionList)
        {
            string sql = $@"INSERT INTO Question(type_id, question_content, question_picture)
                            VALUES('{questionList.QuestionData.type_id}', '{questionList.QuestionData.question_content}',
                            '{questionList.QuestionData.question_picture}')";
            // 先執行當前題目內容
            using var conn = new SqlConnection(cnstr);
            conn.Execute(sql);

            InsertOption(questionList);
        }
        
        // 儲存選項
        public void InsertOption(QuestionList questionList)
        {
            int question_id = GetQuestionId(questionList);
            StringBuilder stringBuilder = new StringBuilder();
            string sql = $@"INSERT INTO Answer(question_id, option_content, question_parse)
                            VALUES({question_id},'{questionList.AnswerData.option_content}',
                            '{questionList.AnswerData.question_parse}')";

            // 是非題
            if(questionList.QuestionData.type_id == 1)
            {
                List<string> Options = new List<string>();
                Options = ["是", "否"];
                for (int i = 0; i < 2; i++)
                {
                    sql += $@"INSERT INTO ""Option""(question_id, option_content)
                                            VALUES({question_id}, '{Options[i]}')";
                }
            }
            
            // 選擇題
            else if(questionList.QuestionData.type_id == 2)
            {
                for(int i = 0; i < 4; i++)
                {
                    sql += $@"INSERT INTO ""Option""(question_id, option_content)
                                            VALUES({question_id}, '{questionList.Options[i]}')";
                }
            }
            
            // 問答題
            else if (questionList.QuestionData.type_id == 3)
            {
                
            }

            // 執行Sql
            using var conn = new SqlConnection(cnstr);
            conn.Execute(sql);
        }

        // 獲得題目id
        public int GetQuestionId(QuestionList question){
            string sql = $@" SELECT question_id FROM Question WHERE question_content = '{question.QuestionData.question_content}' ";
            using var conn = new SqlConnection(cnstr);
            return conn.QueryFirstOrDefault<int>(sql);
        }
    }
}
