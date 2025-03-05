using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace reportBuilder.classes
{
    public class WorkspaceConfigManager
    {
        public WorkspaceConfig LoadConfig(string configFilePath)
        {
            if (File.Exists(configFilePath))
            {
                string json = File.ReadAllText(configFilePath);
                return JsonConvert.DeserializeObject<WorkspaceConfig>(json);
            }
            return null;
        }
    }
}
