using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
	//For GitHub check
    enum blockProperties { Normal, Proj, smVid, smTip }
    

    struct codeBlock
    {
        public string block;
        public blockProperties property;
    }

    class Structure
    {
        public readonly string[] projects = new string[] { "Химки", "Клязьма", "Лопатино", "Сходня", "Опалиха", "Большая Опалиха", "Сабурово", "Чехов", "Центр О", "Проект Коттедж Б", "Ивакино", "Аксаково" };
        public readonly string[] vidSmt = new string[] { "Основная", "Дополнительная" };
        public readonly string[] tipSmt = new string[] { "Смета подрядчика", "Смета заказчика", "Смета внешнего подрядчика", "Смета субподряда" };
        

        public readonly string[] codeProj = new string[] { "ХИ", "КЛ", "ЛО", "СХ", "О2", "О3", "СА", "ЧЕ", "ЦО", "ТЗ", "ИП", "АК" };
        
        
    }

    class DivideProcessor
    {
        private const char gSep = '_';
        private const char cSep = '-';
        private const string ext = ".xml";

        public string sourceString { get; private set; }
        public string Code { get; private set; }
        public string Obj { get; private set; }
        public string Name { get; private set; }
        public codeBlock[] cbBlocks { get; private set; }

        public DivideProcessor() {}
        public DivideProcessor(string text) { sourceString = text; divide(); }
        public void divide()
        {
            string[] gBlocks = generalDivide();
            Code = gBlocks[0];
            Obj = gBlocks[1];
            Name = gBlocks[2];
            cbBlocks = codeDivide();

            int i = 0;
            foreach (string s in gBlocks)
            {
                i++;
                Console.WriteLine("String {1}: {0}", s, i);
            }
            for (i = 0; i < cbBlocks.Length; i++) Console.WriteLine("CodeBlock[{0}] = {1}", i, cbBlocks[i].block);
        }

        private bool genChk()
        {
            if (suppStr.symbSum(sourceString, gSep) == 2) return true;
            else return false;
        }

        private bool codeChk()
        {
            if (suppStr.symbSum(Code, cSep) > 0) return true;
            else return false;
        }

        private void extClean()
        {
            if (sourceString.IndexOf(ext) > 0) sourceString = sourceString.Remove(sourceString.IndexOf(ext));
        }

        private string[] generalDivide()
        {
            if (genChk())
            {
                extClean();
                return sourceString.Split(gSep);
            }
            else return null;
        }

        private codeBlock[] codeDivide()
        {
            if (codeChk())
            {
                string[] arr = Code.Split(cSep);
                codeBlock[] cb = new codeBlock[arr.Length];
                for (int i = 0; i < cb.Length; i++)
                {
                    cb[i].block = arr[i];
                    cb[i].property = blockProperties.Normal;
                }
                return cb;
            }
            else return null;
        }
    }

}
