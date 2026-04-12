using System;
using System.Collections.Generic;
using SimpleNetMail;
using Xunit;

namespace TlsMailSender.Tests
{
    public class MailSenderTests
    {
        // ── NormalizeThumbprint ─────────────────────────────────────────────────
        [Theory]
        [InlineData("aa:bb-cc dd", "AABBCCDD")]
        [InlineData("  ab cd  ",   "ABCD")]
        [InlineData("",            "")]
        [InlineData("---",         "")]
        [InlineData("AB:CD:EF",    "ABCDEF")]
        [InlineData("ab cd ef",    "ABCDEF")]
        public void NormalizeThumbprint_matches_legacy_string_replace(string input, string expected)
        {
            string legacy = input.Trim()
                .Replace(" ", "")
                .Replace(":", "")
                .Replace("-", "")
                .ToUpperInvariant();

            string actual = MailSender.NormalizeThumbprint(input);

            Assert.Equal(legacy, actual);
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void NormalizeThumbprint_returns_empty_for_over_128_chars()
        {
            string longInput = new string('A', 129);
            Assert.Equal(string.Empty, MailSender.NormalizeThumbprint(longInput));
        }

        // ── ParseRecipients ─────────────────────────────────────────────────────
        [Fact]
        public void ParseRecipients_matches_manual_split_behavior()
        {
            string to = " a@x.com ; b@y.com, c@z.com ";
            string display = " A ; B , ";

            string[] recipients = to.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            string[] displayNames = display.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);

            var expected = new List<(string addr, string dn)>();
            for (int i = 0; i < recipients.Length; i++)
            {
                string addr = recipients[i].Trim();
                string dn = (i < displayNames.Length && !string.IsNullOrWhiteSpace(displayNames[i]))
                    ? displayNames[i].Trim()
                    : null;
                expected.Add((addr, dn));
            }

            var actual = MailSender.ParseRecipients(to, display);
            Assert.Equal(expected.Count, actual.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.Equal(expected[i].addr, actual[i].Address);
                Assert.Equal(expected[i].dn, actual[i].DisplayName);
            }
        }

        [Fact]
        public void ParseRecipients_null_display_name_yields_null_per_entry()
        {
            var list = MailSender.ParseRecipients("a@b.com", null);
            Assert.Single(list);
            Assert.Equal("a@b.com", list[0].Address);
            Assert.Null(list[0].DisplayName);
        }

        [Fact]
        public void ParseRecipients_empty_to_returns_empty_list()
        {
            Assert.Empty(MailSender.ParseRecipients("", null));
            Assert.Empty(MailSender.ParseRecipients("   ", null));
        }
    }
}
