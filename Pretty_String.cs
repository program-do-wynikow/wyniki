using System;
using System.Collections.Generic;
using System.Text;

namespace WIZUALIZACJA_CAT_STREAM
{
    class Pretty_String
    {

        //do dostosowania
        private const int names_max_size = 52;
        public static string Generate(string Names, string Club)
        {
            string temp = Names;
            for (int i = 0; i < (names_max_size - Names.Length + 1); i++) temp += " ";
            temp += Club;

            return temp;
        }
    }
}
