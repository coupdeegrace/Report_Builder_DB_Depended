using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace reportBuilder.classes
{
    internal class errorTypeDictionaryClass
    {
        [Name("Тип ошибки")]
        public string errorType {  get; set; }

        [Name("Типизация")]
        public string specification { get; set; }
    }
}
