// Copyright (c) MingMZ. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using Xunit;
using R1C1Converter;

namespace R1C1Converter.Test
{
    public class ConverterTests
    {
        [Theory]
        [InlineData("A1", 1, 1)]
        [InlineData("Z10", 10, 26)]
        [InlineData("AA100", 100, 27)]
        [InlineData("AZ1000", 1000, 52)]
        [InlineData("BA10000", 10000, 53)]
        [InlineData("ZZ100000", 100000, 702)]
        [InlineData("AAA100000", 100000, 703)]
        [InlineData("A2147483647", 2147483647, 1)]
        [InlineData("FXSHRXW1", 1, 2147483647)]
        public void A1ToR1C1ToA1Test(string a1, int r1, int c1)
        {
            string a2;

            R1C1Converter.ToR1C1(a1, out int r2, out int c2);
            a2 = R1C1Converter.FromR1C1(r2, c2);

            Assert.Equal(r1, r2);
            Assert.Equal(c1, c2);
            Assert.Equal(a1, a2);
        }

        [Theory]
        [InlineData("A1:A1", 1, 1, 1, 1)]
        [InlineData("A1:B2", 1, 1, 2, 2)]
        [InlineData("B1:D3", 1, 2, 3, 4)]
        [InlineData("A1:FXSHRXW2147483647", 1, 1, 2147483647, 2147483647)]
        [InlineData("FXSHRXW2147483647:A1", 2147483647, 2147483647, 1, 1)]
        [InlineData("FXSHRXW2147483647:FXSHRXW2147483647", 2147483647, 2147483647, 2147483647, 2147483647)]
        public void A1RangeToR1C1RangeToA1RangeTest(string a1, int r1, int c1, int r2, int c2)
        {
            string reaultA1;

            R1C1Converter.ToR1C1Range(a1, out int resultR1, out int resultC1, out int resultR2, out int resultC2);
            reaultA1 = R1C1Converter.FromR1C1Range(resultR1, resultC1, resultR2, resultC2);

            Assert.Equal(r1, resultR1);
            Assert.Equal(c1, resultC1);
            Assert.Equal(r2, resultR2);
            Assert.Equal(c2, resultC2);
            Assert.Equal(a1, reaultA1);

        }

        [Theory]
        [InlineData("")]
        [InlineData("A")]
        [InlineData("1")]
        [InlineData(" ")]
        [InlineData(".")]
        public void A1InvalidFormatTest(string a1)
        {
            Assert.Throws<FormatException>(() =>
                R1C1Converter.ToR1C1(a1, out _, out _));
        }

        [Fact]
        public void A1NullTest()
        {
            Assert.Throws<ArgumentNullException>(() =>
                R1C1Converter.ToR1C1(null, out _, out _));
        }

        [Theory]
        [InlineData(0, 1)]
        [InlineData(1, 0)]
        public void R1C1InvalidTest(int r1, int c1)
        {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                R1C1Converter.FromR1C1(r1, c1));
        }

        [Theory]
        [InlineData("A:B2")]
        [InlineData("1:B2")]
        [InlineData("A1:B")]
        [InlineData("A1:2")]
        [InlineData("A1,B2")]
        [InlineData("A1B2")]
        public void A1RangeInvalidFormatTest(string a1)
        {
            Assert.Throws<FormatException>(() =>
                R1C1Converter.ToR1C1Range(a1, out _, out _, out _, out _));
        }
    }
}
