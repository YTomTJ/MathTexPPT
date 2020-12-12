using MathTex.Parser;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MathTexPPT {

    public partial class Editor : Form {

        internal double baseScale = 10.0;
        internal string inputLatex = null;
        internal double outputScale = 2.5;
        internal Image outputImage = null;
        internal string outputSvg = null;

        public Editor() {
            InitializeComponent();

            this.KeyPreview = true;
            this.KeyDown += Editor_KeyDown;
            this.txtInput.KeyDown += TxtInput_KeyDown;
        }

        private void Editor_Load(object sender, EventArgs e) {
            cmbFontSize.SelectedIndex = 4;
            txtInput.Text = inputLatex;
            
            picFormula.Image = null;
            outputImage = null;
            outputSvg = null;

            // Load previous formula
            Preview();
        }

        private void Editor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e) {
            if(e.KeyCode == Keys.Escape) {
                Convert();
            }
        }

        private void TxtInput_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e) {
            if(e.KeyCode == Keys.Enter) {
                if(e.Control) {
                    e.Handled = true;
                    txtOutputInfo.Clear();
                    Preview();
                }
            }
        }

        private void butPreview_Click(object sender, EventArgs e) {
            txtOutputInfo.Clear();
            Preview();
        }

        private void butConvert_Click(object sender, EventArgs e) {
            Convert();
        }

        private void butInfo_Click(object sender, EventArgs e) {
            txtOutputInfo.Text = outputSvg;
        }

        private void Editor_FormClosing(object sender, FormClosingEventArgs e) {
            if(e.CloseReason == CloseReason.UserClosing) {
                e.Cancel = true;
                outputImage = null;
                outputSvg = null;
                this.DialogResult = DialogResult.Cancel;
                this.Hide();
            }
        }

        private void Preview() {
            txtInput.Enabled = false;
            //var ms = Renderer.ConvertFormula(txtInput.Text, out string err);
            var bitmap = Renderer.ConvertFormulaEX(txtInput.Text, out string err, out string svgtext);
            if(err is null) {
                //picFormula.Image = new Bitmap(ms);
                outputSvg = svgtext;
                picFormula.Image = bitmap;
                txtOutputInfo.Text += "\nSucceed.";
            } else {
                txtOutputInfo.Text += "\n" + err;
            }
            txtInput.Enabled = true;
            txtInput.Focus();
        }

        private void Convert() {
            try {
                inputLatex = txtInput.Text;
                outputImage = Renderer.ConvertFormulaEX(inputLatex, out string err, out string svgtext, scale: baseScale, color: Color.White);
                outputSvg = svgtext;
                this.DialogResult = DialogResult.OK;
            } catch {
                outputImage = null;
                outputSvg = null;
                this.DialogResult = DialogResult.None;
            }
            this.Hide();
        }

        internal void SetInfo(string info) {
            txtOutputInfo.Text = info;
        }

        private void numScale_ValueChanged(object sender, EventArgs e) {
            outputScale = (double)numScale.Value;
        }
    }
}
