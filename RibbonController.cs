using System;
using CirLat.Properties;
using System.Drawing;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace CirLat {
    [ComVisible(true)]
    public class RibbonController : Microsoft.Office.Core.IRibbonExtensibility {
        private string[] lat = { "dž", "nj", "lj", "Dž", "Nj", "Lj", "DŽ", "NJ", "LJ", "A", "B", "V", "G", "D", "Đ", "E", "Ž", "Z", "I", "J", "K", "L", "M", "N", "O", "P", "R", "S", "T", "Ć", "U", "F", "H", "C", "Č", "Š", "a", "b", "v", "g", "d", "đ", "e", "ž", "z", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "ć", "u", "f", "h", "c", "č", "š" };
        private string[] cir = { "џ", "њ", "љ", "Џ", "Њ", "Љ", "Џ", "Њ", "Љ", "А", "Б", "В", "Г", "Д", "Ђ", "Е", "Ж", "З", "И", "Ј", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "Ћ", "У", "Ф", "Х", "Ц", "Ч", "Ш", "а", "б", "в", "г", "д", "ђ", "е", "ж", "з", "и", "ј", "к", "л", "м", "н", "о", "п", "р", "с", "т", "ћ", "у", "ф", "х", "ц", "ч", "ш" };
        public string GetCustomUI(string ribbonID) =>
            @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                        <ribbon>
                           <tabs>
                                <tab id='preslovljavanje_tab' label='Preslovljavanje'>
                                    <group id='preslovljavanje_group' label='Preslovljavanje'>
                                        <button id='button1' label='Pretvori u ćirilicu' size='large' getImage='GetCirImg' onAction='ToCir'/>
                                        <button id='button2' label='Pretvori u latinicu' size='large' getImage='GetLatImg' onAction='ToLat'/>
                                    </group>
                                </tab>
                            </tabs>
                        </ribbon>
                    </customUI>";

        public void FindReplace(string[] find, string[] replace) {
            var app = Globals.ThisAddIn.Application;
            var document = app.ActiveDocument;
            var window = document.ActiveWindow;
            var recorder = app.UndoRecord;
            bool selected = window.Selection.Type == Word.WdSelectionType.wdSelectionNormal;
            bool footnotes = !selected && document.Footnotes.Count > 0;
            bool endnotes = !selected && document.Endnotes.Count > 0;
            var range1 = selected ? window.Selection.Range : document.Content;
            var range2 = footnotes ? document.Footnotes[1].Range : null;
            var range3 = endnotes ? document.Endnotes[1].Range : null;
            recorder.StartCustomRecord("Preslovljavanje");
            document.Bookmarks.Add("bugfix", document.Range(0, 0));
            document.Bookmarks["bugfix"].Delete();
            range1.Find.ClearFormatting();
            if (footnotes) {
                range2.WholeStory();
                range2.Find.ClearFormatting();
            }
            if (endnotes) {
                range3.WholeStory();
                range3.Find.ClearFormatting();
            }
            var progress = new Progress((Convert.ToInt32(footnotes) + Convert.ToInt32(endnotes) + 1) * find.Length);
            progress.Show();
            for (int i = 0; i < find.Length; i++) {
                range1.Find.Execute(FindText: find[i], ReplaceWith: replace[i], Replace: Word.WdReplace.wdReplaceAll, MatchCase: true);
                progress.nextStep();
                if (footnotes) {
                    range2.Find.Execute(FindText: find[i], ReplaceWith: replace[i], Replace: Word.WdReplace.wdReplaceAll, MatchCase: true);
                    progress.nextStep();
                }
                if (endnotes) {
                    range3.Find.Execute(FindText: find[i], ReplaceWith: replace[i], Replace: Word.WdReplace.wdReplaceAll, MatchCase: true);
                    progress.nextStep();
                }
            }
            recorder.EndCustomRecord();
            progress.Close();
        }

        public void ToCir(Microsoft.Office.Core.IRibbonControl control) => FindReplace(lat, cir);
        public void ToLat(Microsoft.Office.Core.IRibbonControl control) => FindReplace(cir, lat);
        public Bitmap GetCirImg(Microsoft.Office.Core.IRibbonControl control) => Resources.cir;
        public Bitmap GetLatImg(Microsoft.Office.Core.IRibbonControl control) => Resources.lat;
    }
}