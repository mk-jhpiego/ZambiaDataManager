using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
