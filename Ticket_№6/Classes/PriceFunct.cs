using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ticket__6.Classes
{
    internal class PriceFunct
    {
        public static string ImagePath = String.Empty;
        public static bool IsUpdate = false;


        public static int PrintVariant = 0;

        public static bool CheckAllSpace(string Text) //Проверка на все пробелы или пустоту
        {

            if (Text == null || Text.Length == 0) return true;
            for (int i = 0; i < Text.Length; i++) if (!char.IsWhiteSpace(Text[i])) return false;
            return true;

        }
    }
}
