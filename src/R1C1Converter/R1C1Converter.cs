// Copyright (c) MingMZ. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Runtime.CompilerServices;

namespace R1C1Converter
{
    /// <summary>
    /// Convert between A1 and R1C1 reference styles used in Excel document
    /// </summary>
    public class R1C1Converter
    {
        // a valid A1 reference have minimal 2 characters, e.g. "A1", "B2", "C3"
        private const int A1MinLength = 2;

        // using (32-bit)int maximum values
        // - converted to R1 component takes 7 character
        // - converted to C1 component takes 10 Characters
        // so the maximum length of A1 style reference will occupy 17 chars
        private const int A1MaxLength = 17;

        private const char ColumnChar = ':';

        #region conversion from A1 to R1C1 reference style

        /// <summary>
        /// Convert alphabetic component of A1 reference to numeric column index of R1C1
        /// </summary>
        /// <param name="pointer">Pointer to char array to read characters from</param>
        /// <param name="length">Number of chars to read</param>
        /// <returns>R1 value represent column number</returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe int ToC1(char* pointer, int length)
        {
            var value = 0;

            for (int i = 0; i < length; i++)
            {
                value *= 26;
                value += (*pointer - 65) % 26 + 1;
                pointer++;
            }

            return value;
        }

        /// <summary>
        /// Convert numeric component of A1 reference to numeric row index of R1C1
        /// </summary>
        /// <param name="pointer">Pointer to char array to read characters from</param>
        /// <param name="length">Number of chars to read</param>
        /// <returns>C1 value represent row number</returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe int ToR1(char* pointer, int length)
        {
            var value = 0;

            for (int i = 0; i < length; i++)
            {
                value *= 10;
                value += *pointer - 48;
                pointer++;
            }

            return value;
        }

        /// <summary>
        /// Verify A1 reference is valid.
        /// If valid return the position of first numeric character.
        /// Return 0 if not a valid A1 reference.
        /// </summary>
        /// <param name="pointer">Pointer to char array to read characters from</param>
        /// <param name="length">Number of chars to read</param>
        /// <returns>Position of first numeric character</returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe int IsValidA1Reference(char* pointer, int length)
        {
            var position = 0;

            // the first character must be alpha character
            if (65 <= *pointer && *pointer <= 90)
            {
                pointer++;
                for (int i = 1; i < length; i++)
                {
                    // check A (65) to Z (90) until we find numeric chars
                    if (position == 0 && (*pointer < 65 || 90 < *pointer))
                    {
                        // mark the position when character in buffer is no longer in A-Z range
                        // we expect numeric characters 0-9 from this point on
                        position = i;
                    }

                    // check 0 (48) to 9 (57)
                    if (position > 0 && (*pointer < 48 || 57 < *pointer))
                    {
                        // indicate result is not a valid A1 reference
                        position = 0;
                        break;
                    }

                    pointer++;
                }
            }
            return position;
        }

        /// <summary>
        /// Convert from A1 to R1C1 reference style
        /// </summary>
        /// <param name="pointer">Pointer to char array to read characters from</param>
        /// <param name="length">Number of chars to read</param>
        /// <param name="r1">Row index, 1-based</param>
        /// <param name="c1">Column index, 1-based</param>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe void ToR1C1(char* pointer, int length, out int r1, out int c1)
        {
            int position = IsValidA1Reference(pointer, length);

            if (position == 0)
            {
                throw new FormatException("Not valid A1 reference style");
            }

            c1 = ToC1(pointer, position);

            r1 = ToR1(pointer + position, length - position);
        }

        /// <summary>
        /// Convert from A1 to R1C1 reference style
        /// </summary>
        /// <param name="a1">Array of characters to read characters from</param>
        /// <param name="startIndex">Start read position</param>
        /// <param name="length">Characters to read</param>
        /// <param name="r1">Row index, 1-based</param>
        /// <param name="c1">Column index, 1-based</param>
        public static void ToR1C1(char[] a1, int startIndex, int length, out int r1, out int c1)
        {
            if (a1 == null)
            {
                throw new ArgumentNullException();
            }

            unsafe
            {
                fixed (char* pointer = &a1[startIndex])
                {
                    ToR1C1(pointer, length, out r1, out c1);
                }
            }
        }

        /// <summary>
        /// Convert Excel cell position from A1 to R1C1 reference style
        /// </summary>
        /// <param name="a1">A1 reference style</param>
        /// <param name="r1">Row number</param>
        /// <param name="c1">Column number</param>
        public static void ToR1C1(string a1, out int r1, out int c1)
        {
            if (a1 == null)
            {
                throw new ArgumentNullException();
            }

            unsafe
            {
                fixed (char* pointer = a1)
                {
                    ToR1C1(pointer, a1.Length, out r1, out c1);
                }
            }
        }

        /// <summary>
        /// Convert Excel cell range from A1 to R1C1 reference style
        /// </summary>
        /// <param name="pointer">Pointer to char array to read characters from</param>
        /// <param name="length">Number of chars to read</param>
        /// <param name="r1">Begin of row number</param>
        /// <param name="c1">Begin of column number</param>
        /// <param name="r2">End of row number</param>
        /// <param name="c2">End of column number</param>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe void ToR1C1Range(char* pointer, int length, out int r1, out int c1, out int r2, out int c2)
        {
            var localPtr = pointer;

            var separatorPos = 0;

            for (int i = 0; i < length; i++)
            {
                if (*localPtr == ':')
                {
                    separatorPos = i;
                    break;
                }
                localPtr++;
            }

            if (separatorPos < A1MinLength || (length - A1MinLength) < separatorPos)
            {
                throw new FormatException("Not valid A1 reference style range");
            }

            ToR1C1(pointer, separatorPos, out r1, out c1);

            // move past separator
            pointer += separatorPos + 1;
            ToR1C1(pointer, length - separatorPos - 1, out r2, out c2);
        }

        /// <summary>
        /// Convert Excel cell range from A1 to R1C1 reference style
        /// </summary>
        /// <param name="a1">Excel cell range in A1 reference style</param>
        /// <param name="startIndex">Start read position</param>
        /// <param name="length">Number of chars to read</param>
        /// <param name="r1">Begin of row number</param>
        /// <param name="c1">Begin of column number</param>
        /// <param name="r2">End of row number</param>
        /// <param name="c2">End of column number</param>
        public static void ToR1C1Range(char[] a1, int startIndex, int length, out int r1, out int c1, out int r2, out int c2)
        {
            if (a1 == null)
            {
                throw new ArgumentNullException();
            }

            unsafe
            {
                fixed (char* pointer = &a1[startIndex])
                {
                    ToR1C1Range(pointer, length, out r1, out c1, out r2, out c2);
                }
            }
        }

        /// <summary>
        /// Convert Excel cell range from A1 to R1C1 reference style
        /// </summary>
        /// <param name="a1">Excel cell range in A1 reference style</param>
        /// <param name="r1">Begin of row number</param>
        /// <param name="c1">Begin of column number</param>
        /// <param name="r2">End of row number</param>
        /// <param name="c2">End of column number</param>
        public static void ToR1C1Range(string a1, out int r1, out int c1, out int r2, out int c2)
        {
            if (a1 == null)
            {
                throw new ArgumentNullException();
            }

            unsafe
            {
                fixed (char* pointer = a1)
                {
                    ToR1C1Range(pointer, a1.Length, out r1, out c1, out r2, out c2);
                }
            }
        }

        #endregion

        #region conversion from R1C1 to A1 reference style

        /// <summary>
        /// Reverse elements in the array
        /// </summary>
        /// <param name="pointer">Pointer to the beginning of array</param>
        /// <param name="length">Number of elements to reverse</param>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe void Reverse(char* pointer, int length)
        {
            char temp = '\0';

            var reverse = pointer + length - 1;
            while (pointer < reverse)
            {
                temp = *pointer;
                *pointer = *reverse;
                *reverse = temp;

                pointer++;
                reverse--;
            }
        }

        /// <summary>
        /// Convert R1C1 column number to A1 reference style
        /// </summary>
        /// <param name="value">Column number</param>
        /// <param name="pointer">char array to write to</param>
        /// <param name="length">Maximum number of chars to write</param>
        /// <returns>Number of chars written</returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe int FromC1(int value, char* pointer, int length)
        {
            int written = 0;
            var localPtr = pointer;

            while (value > 0 && written <= length)
            {
                value--;
                *localPtr = (char)(value % 26 + 65);
                value = value / 26;

                localPtr++;
                written++;
            }

            Reverse(pointer, written);

            return written;
        }

        /// <summary>
        /// Convert R1C1 row number to A1 reference style
        /// </summary>
        /// <param name="value">Row number</param>
        /// <param name="pointer">Char array to write to</param>
        /// <param name="length">Maximum number of chars to write</param>
        /// <returns>Number of chars written</returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe int FromR1(int value, char* pointer, int length)
        {
            int written = 0;
            var localPtr = pointer;

            while (value > 0 && written <= length)
            {
                *localPtr = (char)(value % 10 + 48);
                value = value / 10;

                localPtr++;
                written++;
            }

            Reverse(pointer, written);

            return written;
        }

        /// <summary>
        /// Convert R1C1 to A1 reference style
        /// </summary>
        /// <param name="r">Row number</param>
        /// <param name="c">Column number</param>
        /// <param name="pointer">Char array to write to</param>
        /// <param name="length">Maximum number of chars to write</param>
        /// <returns>Number of chars written</returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe int FromR1C1(int r, int c, char* pointer, int length)
        {
            if (r < 1 || c < 1)
            {
                throw new ArgumentOutOfRangeException("R1C1 value cannot be less than 1");
            }

            var written = FromC1(c, pointer, length);
            written += FromR1(r, pointer + written, length - written);

            return written;
        }

        /// <summary>
        /// Convert R1C1 to A1 reference style
        /// </summary>
        /// <param name="r">Row number</param>
        /// <param name="c">Column number</param>
        /// <param name="a1">Array of characters to write A1 reference to</param>
        /// <param name="startIndex">Starting position to write</param>
        /// <param name="length">Maximum number of chars to write</param>
        /// <returns>Number of chars written</returns>
        public static int FromR1C1(int r, int c, char[] a1, int startIndex, int length)
        {
            if (a1 == null)
            {
                throw new ArgumentNullException();
            }

            unsafe
            {
                fixed (char* pointer = &a1[startIndex])
                {
                    return FromR1C1(r, c, pointer, length);
                }
            }
        }

        /// <summary>
        /// Convert R1C1 to A1 reference style
        /// </summary>
        /// <param name="r">Row number</param>
        /// <param name="c">Column number</param>
        /// <returns>A1 reference style</returns>
        public static string FromR1C1(int r, int c)
        {
            var a1 = new char[A1MaxLength];
            var written = 0;

            unsafe
            {
                fixed (char* pointer = &a1[0])
                {
                    written = FromR1C1(r, c, pointer, a1.Length);
                }
            }

            return new string(a1, 0, written);
        }

        /// <summary>
        /// Convert R1C1 range to A1 reference style
        /// </summary>
        /// <param name="r1">Begin of row number</param>
        /// <param name="c1">Begin of column number</param>
        /// <param name="r2">End of row number</param>
        /// <param name="c2">End of column number</param>
        /// <param name="pointer">Char array to write A1 reference to</param>
        /// <param name="length">Maximum number of chars to write</param>
        /// <returns>Number of chars written</returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static unsafe int FromR1C1Range(int r1, int c1, int r2, int c2, char* pointer, int length)
        {
            var written = FromR1C1(r1, c1, pointer, length);

            pointer += written;
            *pointer = ColumnChar;
            written++;

            written += FromR1C1(r2, c2, pointer + 1, length - written);
            return written;
        }

        /// <summary>
        /// Convert R1C1 range to A1 reference style
        /// </summary>
        /// <param name="r1">Begin of row number</param>
        /// <param name="c1">Begin of column number</param>
        /// <param name="r2">End of row number</param>
        /// <param name="c2">End of column number</param>
        /// <param name="a1">Array of characters to write A1 reference to</param>
        /// <param name="startIndex">Starting position to write</param>
        /// <param name="length">Maximum number of chars to write</param>
        /// <returns>Number of chars written</returns>
        public static int FromR1C1Range(int r1, int c1, int r2, int c2, char[] a1, int startIndex, int length)
        {
            if (a1 == null)
            {
                throw new ArgumentNullException();
            }

            unsafe
            {
                fixed (char* pointer = &a1[startIndex])
                {
                    return FromR1C1Range(r1, c1, r2, c2, pointer, length);
                }
            }
        }

        /// <summary>
        /// Convert R1C1 range to A1 reference style
        /// </summary>
        /// <param name="r1">Begin of row number</param>
        /// <param name="c1">Begin of column number</param>
        /// <param name="r2">End of row number</param>
        /// <param name="c2">End of column number</param>
        /// <returns>A1 reference style range</returns>
        public static string FromR1C1Range(int r1, int c1, int r2, int c2)
        {
            var a1 = new char[A1MaxLength * 2 + 1];
            var written = 0;

            unsafe
            {
                fixed (char* pointer = &a1[0])
                {
                    written = FromR1C1Range(r1, c1, r2, c2, pointer, a1.Length);
                }
            }

            return new string(a1, 0, written);
        }

        #endregion
    }
}
