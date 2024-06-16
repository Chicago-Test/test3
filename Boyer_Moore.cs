using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace VBAStreamDecompress
{
    // https://en.wikipedia.org/wiki/Boyer%E2%80%93Moore_string-search_algorithm
    // Boyer–Moore string-search algorithm
    // Converted JAVA sample
    internal class Boyer_Moore
    {
        /**
    * Returns the index within this string of the first occurrence of the
    * specified substring. If it is not a substring, return -1.
    *
    * There is no Galil because it only generates one match.
    *
    * @param haystack The string to be scanned
    * @param needle The target string to search
    * @return The start index of the substring
    */
        public static int indexOf(byte[] haystack, byte[] needle)
        {
            if (needle.Length == 0)
            {
                return 0;
            }
            int[] charTable = makeCharTable(needle);
            int[] offsetTable = makeOffsetTable(needle);
            for (int i = needle.Length - 1, j; i < haystack.Length;)
            {
                for (j = needle.Length - 1; needle[j] == haystack[i]; --i, --j)
                {
                    if (j == 0)
                    {
                        return i;
                    }
                }
                // i += needle.Length - j; // For naive method
                i += Math.Max(offsetTable[needle.Length - 1 - j], charTable[haystack[i]]);
            }
            return -1;
        }

        /**
         * Makes the jump table based on the mismatched character information.
         * (bad-character rule)
         */
        private static int[] makeCharTable(byte[] needle)
        {
            int ALPHABET_SIZE = char.MaxValue + 1;
            int[] table = new int[ALPHABET_SIZE];
            for (int i = 0; i < table.Length; ++i)
            {
                table[i] = needle.Length;
            }
            for (int i = 0; i < needle.Length; ++i)
            {
                table[needle[i]] = needle.Length - 1 - i;
            }
            return table;
        }

        /**
         * Makes the jump table based on the scan offset which mismatch occurs.
         * (good suffix rule)
         */
        private static int[] makeOffsetTable(byte[] needle)
        {
            int[] table = new int[needle.Length];
            int lastPrefixPosition = needle.Length;
            for (int i = needle.Length; i > 0; --i)
            {
                if (isPrefix(needle, i))
                {
                    lastPrefixPosition = i;
                }
                table[needle.Length - i] = lastPrefixPosition - i + needle.Length;
            }
            for (int i = 0; i < needle.Length - 1; ++i)
            {
                int slen = suffixLength(needle, i);
                table[slen] = needle.Length - 1 - i + slen;
            }
            return table;
        }

        /**
         * Is needle[p:end] a prefix of needle?
         */
        private static bool isPrefix(byte[] needle, int p)
        {
            for (int i = p, j = 0; i < needle.Length; ++i, ++j)
            {
                if (needle[i] != needle[j])
                {
                    return false;
                }
            }
            return true;
        }

        /**
         * Returns the maximum length of the substring ends at p and is a suffix.
         * (good-suffix rule)
         */
        private static int suffixLength(byte[] needle, int p)
        {
            int len = 0;
            for (int i = p, j = needle.Length - 1;
                     i >= 0 && needle[i] == needle[j]; --i, --j)
            {
                len += 1;
            }
            return len;
        }
    }
}
