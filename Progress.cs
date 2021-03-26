using System.Windows.Forms;

namespace CirLat {
    public partial class Progress : Form {
        public Progress(int steps) {
            InitializeComponent();
            CenterToParent();
            progressBar1.Maximum = steps;
        }
        public void nextStep() {
            progressBar1.Value++;
            label1.Text = (progressBar1.Value * 100 / progressBar1.Maximum) + "%";
            progressBar1.Refresh();
            label1.Refresh();
        }
    }
}
