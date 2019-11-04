using System;

namespace ASTA.Classes
{
  public static class CompareTwoStrings
    {
         public static int Compare(string a, string b)
        {
            string[] words = { a, b };
            Array.Sort(words);

            if (words[0] != a)
            {
                return 1;
            }
            else if (a == b)
            {
                return 0;
            }
            else
            {
                return -1;
            }
        }
    }
}
