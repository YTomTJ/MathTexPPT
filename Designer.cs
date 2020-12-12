using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace MathTexPPT {

    public partial class Designer {

        private static Editor editor;
        private readonly string IDHeader = "##_MathTexPPT";
        private readonly string IDTail = "TPPxeThtaM_##";
        private static ulong IDIndex = 0;

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
            var id = GetId(shape);
            if(id is null) {
                return;
            }
            editor.Text = $"Editor - {id}";
            editor.inputLatex = GetFormula(shape);
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
                SetFormula(shape, formula);
                shape.ScaleHeight((float)(editor.outputScale / editor.baseScale), Microsoft.Office.Core.MsoTriState.msoTrue);
                shape.ScaleWidth((float)(editor.outputScale / editor.baseScale), Microsoft.Office.Core.MsoTriState.msoTrue);
                return shape;
            } else {
                MessageBox.Show("Paste output image failed.");
            }
            return null;
        }

        #region MathTexPPT Tools

        private string GenerateId() {
            return $"{IDHeader};{IDIndex++};{DateTime.Now.ToString("yyyyMMddHHmmss")}:\\sqrt{{A}}:{IDTail}";
        }

        private bool CheckDescr(PPT.Shape shape) {
            return shape.Title != null
                && shape.Title.Contains(IDHeader)
                && shape.Title.Contains(IDTail);
        }

        private string GetId(PPT.Shape shape) {
            if(CheckDescr(shape)) {
                int start = shape.Title.IndexOf(IDHeader);
                int split = shape.Title.IndexOf(':', start);
                return shape.Title.Substring(start, split - start);
            }
            return null;
        }

        private string GetFormula(PPT.Shape shape) {
            if(CheckDescr(shape)) {
                int start = shape.Title.IndexOf(IDHeader);
                int split = shape.Title.IndexOf(':', start) + 1;
                int end = shape.Title.IndexOf(IDTail) - 1;
                return shape.Title.Substring(split, end - split);
            }
            return null;
        }

        private string SetFormula(PPT.Shape shape, string formula = null) {
            if(!CheckDescr(shape)) {
                shape.Title = GenerateId();
            }
            if(formula != null) {
                int start = shape.Title.IndexOf(IDHeader);
                int split = shape.Title.IndexOf(':', start);
                int end = shape.Title.IndexOf(IDTail) + IDTail.Length;
                shape.Title = shape.Title.Substring(0, start)
                    + $"{shape.Title.Substring(start, split - start)}:{formula}:{IDTail}"
                    + shape.Title.Substring(end);
            }
            return shape.Title;
        }

        #endregion MathTexPPT Tools

    }
}
