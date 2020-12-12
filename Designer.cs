using Microsoft.Office.Tools.Ribbon;
using System.Drawing;
using System.Windows.Forms;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace MathTexPPT {

    public partial class Designer {

        private static Editor editor;

        private void Designer_Load(object sender, RibbonUIEventArgs e) {
            editor = new Editor();
        }

        private void butInsert_Click(object sender, RibbonControlEventArgs e) {
            OpenEdit();
        }

        private void butEdit_Click(object sender, RibbonControlEventArgs e) {
            // Check selection
            try {
                var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if(sel != null && sel.ShapeRange != null && sel.ShapeRange.Count > 0) {
                    if(sel.ShapeRange.Count == 1 && sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture) {
                        OpenEdit(sel.ShapeRange[1]);
                    } else {
                        MessageBox.Show("You didn't select a valid formula picture.");
                    }
                }
            } catch {
                MessageBox.Show("You should select a formula picture.");
            }
        }

        private void OpenEdit() {

            // Load formula latex.
            editor.Text = $"Editor - Temproary";
            editor.inputLatex = "";

            // Edit it and retrieve output image.
            if(editor.ShowDialog() == DialogResult.OK) {
                if(editor.outputImage != null) {
                    PastePicture(editor.outputImage, editor.inputLatex);
                }
            }
        }

        private void OpenEdit(PPT.Shape shape) {
            // Load formula latex.
            var id = ShapeManager.GetId(shape);
            if(id is null) {
                return;
            }
            editor.Text = $"Editor - {id}";
            editor.inputLatex = ShapeManager.GetFormula(shape);
            editor.SetInfo($">>> {id}\n");
            // Edit it and retrieve output image.
            if(editor.ShowDialog() == DialogResult.OK) {
                if(editor.outputImage != null) {
                    var left = shape.Left;
                    var top = shape.Top;
                    shape.Delete();
                    var sh = PastePicture(editor.outputImage, editor.inputLatex);
                    sh.Left = left;
                    sh.Top = top;
                }
            }
        }

        private PPT.Shape PastePicture(Image image, string formula = null) {
            PPT.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var c0 = slide.Shapes.Count;
            Clipboard.SetDataObject(image);
            slide.Shapes.Paste();
            if(slide.Shapes.Count == c0 + 1) {
                var shape = slide.Shapes[slide.Shapes.Count];
                ShapeManager.SetFormula(shape, formula);
                shape.ScaleWidth((float)(editor.outputScale / editor.baseScale), Microsoft.Office.Core.MsoTriState.msoTrue);
                shape.ScaleHeight((float)(editor.outputScale / editor.baseScale), Microsoft.Office.Core.MsoTriState.msoTrue);
                return shape;
            } else {
                MessageBox.Show("Paste output image failed.");
            }
            return null;
        }
    }
}
