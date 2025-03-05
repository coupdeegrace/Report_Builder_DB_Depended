using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace reportBuilder.classes
{
    public class WorkspaceConfig
    {
        public string TableName { get; set; }
        public string SpreadsheetId { get; set; }
        public string Range { get; set; }
        public string UpdateRange {  get; set; }
        public string CredentialsPath {  get; set; }
    }
}
