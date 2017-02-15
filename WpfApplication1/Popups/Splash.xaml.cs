using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ZambiaDataManager.Popups
{
    /// <summary>
    /// Interaction logic for SplashScreen.xaml
    /// </summary>
    public partial class Splash : Window, IDisplayProgress
    {
        public Splash()
        {
            InitializeComponent();
        }

        int progressStep = 10;
        int progressStepsRemaining = 100;
        public void PerformProgressStep(string message = "")
        {
            //we reduce time remaining
            progressStepsRemaining -= progressStep;
            progressStepsRemaining = progressStepsRemaining < 0 ? 0 : progressStepsRemaining;
            lblMainMessage.Content = message;
            //lblMainMessage.SetText(message);
            prgMain.SetValue(Convert.ToInt32((100 - progressStepsRemaining) * 1.0));
        }

        public void MarkStartOfMultipleSteps(int stepsToExpect)
        {
            progressStep = Convert.ToInt32(progressStepsRemaining * 1.0 / stepsToExpect);
            prgMain.SetStepValue(progressStep);
        }

        int _totalSubProgressSteps = 0;
        int _totalSubSteps = 0;
        public void ResetSubProgressIndicator(int stepsToExpect)
        {
            _totalSubProgressSteps = stepsToExpect;
            _totalSubSteps = stepsToExpect;
        }

        public void PerformSubProgressStep()
        {
            //we reduce time remaining
            _totalSubProgressSteps--;
            //lblSubMessage.SetText(_totalSubProgressSteps.ToString());
            //prgSubSteps.SetValue(Convert.ToInt32((_totalSubSteps - _totalSubProgressSteps) * 100.0 / _totalSubSteps));
        }
    }
}
