using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Electre_Customize_DotNet.Helpers.CableList
{
    /// <summary>
    /// Cable group-code rules shared by the two loom report generators (with and without the HAL
    /// template): which group codes are invalid, and the order loom wires are listed in.
    /// The patterns are held as compiled instances because they run once per wire entry, per selected
    /// loom - the static Regex.IsMatch/Match/Split calls they replace already cached internally, but a
    /// dedicated compiled instance skips that lookup and JIT-compiles the matching. Same patterns, same
    /// matching behaviour.
    /// </summary>
    internal static class LoomGroupRules
    {
        private static readonly Regex StartsWithLetter = new(@"^[A-Za-z]", RegexOptions.Compiled);
        private static readonly Regex ValidGroupPattern = new(@"^[\dA-Za-z_-]+$", RegexOptions.Compiled);
        private static readonly Regex LeadingNumber = new(@"^(\d+)", RegexOptions.Compiled);
        private static readonly Regex NumberAndAlpha = new(@"^(\d+)([A-Za-z]*)", RegexOptions.Compiled);
        private static readonly Regex SplitDelims = new(@"[-_]", RegexOptions.Compiled);

        /// <summary>
        /// Wires whose (non-empty) group code is invalid: first those starting with a letter, then those
        /// not matching the allowed pattern (digits, letters, '_' and '-') - the order the original
        /// logged them in. Returns (Group, WireCode) pairs.
        /// </summary>
        public static List<(string Group, string WireCode)> FindIncorrectGroups<T>(
            IEnumerable<T> wires, Func<T, string> group, Func<T, string> wireCode)
        {
            var startsWithLetter = wires
                .Where(x => !string.IsNullOrEmpty(group(x)) && StartsWithLetter.IsMatch(group(x)))
                .Select(x => (Group: group(x), WireCode: wireCode(x)))
                .ToList();

            var otherInvalid = wires
                .Where(x => !string.IsNullOrEmpty(group(x)) && !ValidGroupPattern.IsMatch(group(x)))
                .Select(x => (Group: group(x), WireCode: wireCode(x)))
                .ToList();

            return startsWithLetter.Concat(otherInvalid).ToList();
        }

        /// <summary>
        /// Drops wires whose group code starts with a letter, then orders the rest by: leading group
        /// number, the letters after it (case-insensitive), the sub-number after '_' or '-', and finally
        /// the wire code. A null group code throws, exactly as the inline version did.
        /// </summary>
        public static List<T> FilterAndSort<T>(IEnumerable<T> wires, Func<T, string> group, Func<T, string> wireCode)
        {
            return wires
                .Where(x => !StartsWithLetter.IsMatch(group(x))) // Exclude groups starting with a letter
                .OrderBy(x =>
                {
                    // Extract main number from the beginning of the string
                    var match = LeadingNumber.Match(group(x));
                    return match.Success ? int.Parse(match.Groups[1].Value) : int.MaxValue;
                })
                .ThenBy(x =>
                {
                    // Extract alphabetical part after the number (case-insensitive)
                    var match = NumberAndAlpha.Match(group(x));
                    return match.Success ? match.Groups[2].Value.ToLower() : "";
                })
                .ThenBy(x =>
                {
                    // Extract sub-number after _ or - (default to 0 if not found)
                    var parts = SplitDelims.Split(group(x).Replace('_', '-'));
                    return parts.Length > 1 && int.TryParse(parts[1], out int subNumber) ? subNumber : 0;
                })
                .ThenBy(x => wireCode(x))  // Sort by WireCode last
                .ToList();
        }
    }
}
