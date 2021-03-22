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
            bool selected = Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.Selection.Type == Word.WdSelectionType.wdSelectionNormal;
            Word.Range range1 = selected ? Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.Selection.Range : Globals.ThisAddIn.Application.ActiveDocument.Content;
            range1.Find.ClearFormatting();
            for (int i = 0; i < 63; i++)
                range1.Find.Execute(FindText: find[i], ReplaceWith: replace[i], Replace: Word.WdReplace.wdReplaceAll, MatchCase: true);
            if (!selected && Globals.ThisAddIn.Application.ActiveDocument.Footnotes.Count > 0) {
                Word.Range range2 = Globals.ThisAddIn.Application.ActiveDocument.Footnotes[1].Range;
                range2.WholeStory();
                range2.Find.ClearFormatting();
                for (int i = 0; i < 63; i++)
                    range2.Find.Execute(FindText: find[i], ReplaceWith: replace[i], Replace: Word.WdReplace.wdReplaceAll, MatchCase: true);
            }
        }

        public void ToCir(Microsoft.Office.Core.IRibbonControl control) => FindReplace(lat, cir);
        public void ToLat(Microsoft.Office.Core.IRibbonControl control) => FindReplace(cir, lat);
        public Bitmap GetCirImg(Microsoft.Office.Core.IRibbonControl control) => Resources.cir;
        public Bitmap GetLatImg(Microsoft.Office.Core.IRibbonControl control) => Resources.lat;
    }
}