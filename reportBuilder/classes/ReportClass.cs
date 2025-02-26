using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper.Configuration.Attributes;

namespace reportBuilder.classes
{
    public class ReportClass
    {
        [Name ("Номер")]
        public string task_no {  get; set; }
        [Name ("Тип ошибки")]
        public string error_type { get; set; }
        [Name ("Рабочая группа")]
        public string workspace { get; set; }
        [Name ("Краткое описание")]
        public string short_desc {  get; set; }

    }
}
