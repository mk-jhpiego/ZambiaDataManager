using System.Collections.Generic;

namespace ZambiaDataManager
{
    public class PageController
    {
        static PageController _ControllerManager;
        public static PageController Instance
        {
            get
            {
                if (_ControllerManager == null)
                    _ControllerManager = new PageController();
                return _ControllerManager;
            }
        }

        public ProjectName DefaultProjectName { get; set; }
        public Dictionary<string, string> AlternateAgegroups { get; set; }
    }
}
