using CirLat.Properties;
using System.Drawing;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace CirLat {
    [ComVisible(true)]
    public class RibbonController : Microsoft.Office.Core.IRibbonExtensibility {
        private string[] lat = { "A", "B", "V", "G", "D", "Đ", "E", "Ž", "Z", "I", "J", "K", "L", "Lj", "M", "N", "Nj", "O", "P", "R", "S", "T", "Ć", "U", "F", "H", "C", "Č", "Dž", "Š", "a", "b", "v", "g", "d", "đ", "e", "ž", "z", "i", "j", "k", "l", "lj", "m", "n", "nj", "o", "p", "r", "s", "t", "ć", "u", "f", "h", "c", "č", "dž", "š" };
        private string[] cir = { "А", "Б", "В", "Г", "Д", "Ђ", "Е", "Ж", "З", "И", "Ј", "К", "Л", "Љ", "М", "Н", "Њ", "О", "П", "Р", "С", "Т", "Ћ", "У", "Ф", "Х", "Ц", "Ч", "Џ", "Ш", "а", "б", "в", "г", "д", "ђ", "е", "ж", "з", "и", "ј", "к", "л", "љ", "м", "н", "њ", "о", "п", "р", "с", "т", "ћ", "у", "ф", "х", "ц", "ч", "џ", "ш" };
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
        
        public void ToCir(Microsoft.Office.Core.IRibbonControl control) {
            Word.Range range = Globals.ThisAddIn.Application.ActiveDocument.Content;
            range.Find.ClearFormatting();
            for (int i = 0; i < 60; i++)
                range.Find.Execute(FindText: lat[i], ReplaceWith: cir[i], Replace: Word.WdReplace.wdReplaceAll, MatchCase: true);
        }

        public void ToLat(Microsoft.Office.Core.IRibbonControl control) {
            Word.Range range = Globals.ThisAddIn.Application.ActiveDocument.Content;
            range.Find.ClearFormatting();
            for (int i = 0; i < 60; i++)
                range.Find.Execute(FindText: cir[i], ReplaceWith: lat[i], Replace: Word.WdReplace.wdReplaceAll, MatchCase: true);
        }

        public Bitmap GetCirImg(Microsoft.Office.Core.IRibbonControl control) => Resources.cir;
        public Bitmap GetLatImg(Microsoft.Office.Core.IRibbonControl control) => Resources.lat;
    }
}