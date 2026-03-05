using System.Windows;

namespace ScheduleExporter {
    public partial class LoadingWindow : Window {
        public LoadingWindow() {
            InitializeComponent();
        }

        public void UpdateStatus(string status) {
            Dispatcher.Invoke(() => {
                textStatus.Text = status;
            });
        }
    }
}
