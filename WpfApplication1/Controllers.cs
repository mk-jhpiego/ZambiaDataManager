using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApplication1
{
    //public class ControllerManager
    //{
    //    static ControllerManager _ControllerManager;
    //    public static ControllerManager Instance
    //    {
    //        get
    //        {
    //            if (_ControllerManager == null)
    //                _ControllerManager = new ControllerManager();
    //            return _ControllerManager;
    //        }
    //    }
    //    public PageController DefaultController { get; set; }
    //}

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
    }
}
