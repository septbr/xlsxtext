// =============================================================================
// number_format_3.cpp — Excel number format string parser & formatter
//
// Implements ECMA-376 Part 1 §18.8.31 (numFmt / formatCode) semantics.
// Reference: ISO/IEC 29500-1:2016 (Office Open XML File Formats)
//
// Architecture:
//   anonymous namespace — pure math/date helpers, no dependency on impl
//   number_format::impl   — all parsing, analysis, and formatting logic
//   number_format public  — Pimpl wrappers (thin delegation)
// =============================================================================

#include "number_format_3.hpp"

#include <algorithm>
#include <array>
#include <cctype>
#include <charconv>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <numeric>
#include <stdexcept>
#include <vector>

namespace xlsxtext
{

// =============================================================================
// Anonymous namespace — pure helpers
// =============================================================================

namespace
{

    // Floating-point comparison tolerance
    constexpr double float_epsilon = 1e-10;

    // Default Excel column width in characters (for overflow display: "####...")
    constexpr std::size_t default_cell_width = 11;

    // Precomputed powers of 10 for fast lookup (indices 0..15)
    constexpr std::array<double, 16> pow10 = {
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7,
        1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15
    };

    // Safe power-of-10 lookup with fallback to std::pow for large exponents
    inline double pow10_safe(int n) noexcept
    {
        if (n >= 0 && n < static_cast<int>(pow10.size()))
            return pow10[static_cast<size_t>(n)];
        return std::pow(10.0, static_cast<double>(n));
    }

    // ---- elapsed time bracket detection (shared by parse_bracket & tokenize_section) ----

    // Returns the elapsed time type ('h','m','s') if bracket content is all the
    // same h/m/s letter (case-insensitive), or '\0' if not elapsed time.
    // E.g., "h"→'h', "HH"→'h', "mmm"→'m', "ssss"→'s', "Red"→'\0'.
    inline char is_elapsed_bracket(const std::string& s) noexcept
    {
        if (s.empty()) return '\0';
        const char first = static_cast<char>(std::tolower(static_cast<unsigned char>(s[0])));
        if (first != 'h' && first != 'm' && first != 's') return '\0';
        for (size_t k = 1; k < s.size(); ++k)
            if (static_cast<char>(std::tolower(static_cast<unsigned char>(s[k]))) != first)
                return '\0';
        return first;
    }

    // ---- date/time name tables ----

    constexpr std::array<const char*, 12> month_names = {
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    };
    constexpr std::array<const char*, 12> month_abbr = {
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    };
    constexpr std::array<const char*, 7> day_names = {
        "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
    };
    constexpr std::array<const char*, 7> day_abbr = {
        "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"
    };

    // Decomposed date/time components
    struct date_parts
    {
        int year, month, day, hour, minute, second, dow; // dow = day-of-week (0=Sun)
        double fsec; // fractional seconds
    };

    // ---- fast integer-to-string (no heap allocation) ----

    void append_int(std::string& out, int val) noexcept
    {
        char buf[16];
        auto r = std::to_chars(buf, buf + 16, val);
        out.append(buf, r.ptr);
    }

    // Append val with leading zeros to fill at least `width` characters.
    void append_padded(std::string& out, long long val, int width) noexcept
    {
        char buf[32];
        auto r = std::to_chars(buf, buf + sizeof(buf), val);
        const int len = static_cast<int>(r.ptr - buf);
        for (int k = len; k < width; ++k) out += '0';
        out.append(buf, static_cast<size_t>(len));
    }

    // ---- date/time arithmetic (Howard Hinnant's civil calendar algorithms) ----
    int days_from_civil(int y, int m, int d) noexcept
    {
        y -= m <= 2;
        const int era = (y >= 0 ? y : y - 399) / 400;
        const int yoe = static_cast<int>(y - era * 400);
        const int doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1;
        const int doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
        return era * 146097 + doe - 719468;
    }

    // Convert serial day number back to civil date (y,m,d)
    void civil_from_days(int z, int& y, int& m, int& d) noexcept
    {
        z += 719468;
        const int era = (z >= 0 ? z : z - 146096) / 146097;
        const int doe = static_cast<int>(z - era * 146097);
        const int yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
        y = static_cast<int>(yoe + era * 400);
        const int doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
        const int mp = (5 * doy + 2) / 153;
        d = doy - (153 * mp + 2) / 5 + 1;
        m = mp + (mp < 10 ? 3 : -9);
        y += (m <= 2);
    }

    // Compute day-of-week from civil date (0=Sunday .. 6=Saturday).
    // Uses days_from_civil: 1970-01-01 (Thursday, dow=4) maps to day 0.
    inline int calc_dow(int y, int m, int d) noexcept
    {
        const int v = days_from_civil(y, m, d) + 4;
        return v >= 0 ? v % 7 : (v % 7 + 7) % 7;
    }

    // Convert Excel serial date number to date_parts.
    // Handles the 1900 leap-year bug (serial 60 = Feb 29, 1900).
    // Supports both 1900 and 1904 date systems.
    date_parts serial_to_date(double serial, bool date1904) noexcept
    {
        const int whole = static_cast<int>(std::floor(serial));
        const double frac = serial - whole;

        // ECMA-376: 1900 date system treats serial 60 as 1900-02-29 (leap year bug)
        if (!date1904 && whole == 60)
        {
            const int dow = calc_dow(1900, 2, 29);
            const double total_seconds = frac * 86400.0;
            const int hour = static_cast<int>(total_seconds) / 3600;
            const int minute = (static_cast<int>(total_seconds) % 3600) / 60;
            double fsec = total_seconds - hour * 3600.0 - minute * 60.0;
            const int second = static_cast<int>(std::floor(fsec));
            fsec -= second;
            if (fsec >= 1.0) fsec = std::nextafter(1.0, 0.0);
            if (fsec < 0.0)  fsec = 0.0;
            return {1900, 2, 29, hour, minute, second, dow, fsec};
        }

        // 1900 system: days 1..60 have an off-by-one (nonexistent Feb 29, 1900)
        int real_days = whole;
        if (!date1904 && real_days > 60)
            --real_days;

        const int epoch_days = date1904
            ? days_from_civil(1904, 1, 1)   // 1904 epoch
            : days_from_civil(1899, 12, 31); // 1900 epoch (day 1 = 1900-01-01)

        int y, m, d;
        civil_from_days(epoch_days + real_days, y, m, d);

        // Compute dow from the actual civil date — correct for all serials and both date systems
        const int dow = calc_dow(y, m, d);
        const double total_seconds = frac * 86400.0;
        const int hour = static_cast<int>(total_seconds) / 3600;
        const int minute = (static_cast<int>(total_seconds) % 3600) / 60;
        double fsec = total_seconds - hour * 3600.0 - minute * 60.0;
        const int second = static_cast<int>(std::floor(fsec));
        fsec -= second;
        // Clamp to [0, 1) to guard against floating-point error near boundaries
        if (fsec >= 1.0) fsec = std::nextafter(1.0, 0.0);
        if (fsec < 0.0)  fsec = 0.0;

        return {y, m, d, hour, minute, second, dow, fsec};
    }

    // Find the best rational approximation of `frac` with denominator ≤ max_den.
    // Exhaustive search over all denominators 1..max_den — fast enough for
    // typical fraction formats (max_den ≤ 99 for "??/??").
    void best_fraction(double frac, int max_den, long long& best_num, long long& best_den) noexcept
    {
        best_num = 0;
        best_den = 1;
        double best_err = 1.0;

        for (int d = 1; d <= max_den; ++d)
        {
            const int n = static_cast<int>(std::round(frac * d + float_epsilon));
            if (n > d) continue;
            const double err = std::fabs(frac - static_cast<double>(n) / d);
            if (err < best_err - 1e-12) { best_err = err; best_num = n; best_den = d; }
        }

        // Reduce fraction to lowest terms
        if (best_num > 0 && best_den > 1)
        {
            const auto g = std::gcd(static_cast<long long>(best_num), static_cast<long long>(best_den));
            if (g > 1) { best_num /= g; best_den /= g; }
        }
    }

    // "General" format: shortest decimal representation via std::to_chars
    std::string format_number_general(double number)
    {
        char buf[64]{};
        auto r = std::to_chars(buf, buf + sizeof(buf), number);
        return std::string(buf, r.ptr);
    }

} // anonymous namespace

// =============================================================================
// number_format::impl — all internal types and logic
// =============================================================================

struct number_format::impl
{
    // =========================================================================
    // Token types — ECMA-376 §18.8.31 format code components
    // =========================================================================

    enum class token_type : uint8_t
    {
        // ---- structural ----
        literal,          // literal text, quoted strings, escaped chars
        text_placeholder, // @ — replaced with the raw cell text
        skip,             // _x — skip the width of the next character
        fill,             // *x — repeat the next character to fill the cell

        // ---- number placeholders (§18.8.31) ----
        digit_zero,       // 0 — always show digit, pad with leading/trailing zeros
        digit_hash,       // # — show significant digit only, suppress insignificant
        digit_qmark,      // ? — like # but pad with spaces for alignment
        decimal,           // . — decimal point position
        thousands,         // , — thousands separator (or scale divider)
        percent,           // % — multiply by 100, show percent sign
        scientific,        // E+/E- — scientific notation exponent marker

        // ---- date placeholders (§18.17.4.1) ----
        year_2,           // yy
        year_4,           // yyyy+
        month_n,          // m
        month_nn,         // mm
        month_mmm,        // mmm — abbreviated month name
        month_mmmm,       // mmmm — full month name
        month_mmmmm,      // mmmmm+ — first letter of month
        day_d,            // d
        day_dd,           // dd
        day_ddd,          // ddd — abbreviated day name
        day_dddd,         // dddd+ — full day name

        // ---- time placeholders (§18.17.4.1) ----
        hour_h,           // h
        hour_hh,          // hh
        minute_m,         // m (resolved from m/mm ambiguity)
        minute_mm,        // mm (resolved from m/mm ambiguity)
        second_s,         // s
        second_ss,        // ss
        am_pm,            // AM/PM or A/P
        frac_second,      // .0, .00, .000 — fractional seconds

        // ---- elapsed time (§18.8.31) ----
        elapsed_hours,    // [h], [hh], [hhh], ...
        elapsed_minutes,  // [m], [mm], ...
        elapsed_seconds,  // [s], [ss], ...
    };

    // A parsed token — the basic unit after flattening
    struct token
    {
        token_type type = token_type::literal;
        std::string literal;    // literal text content (for literal, skip, fill, am_pm)
        int repeat = 0;         // repeat count (for digit placeholders, frac_second, elapsed time)
        bool exp_plus_sign = true;   // E+ vs E- (for scientific)
        int exp_digits = 0;      // number of exponent digits (for scientific)
        bool exp_upper_case = true; // E vs e (for scientific)
        std::string exp_pattern; // exponent digit pattern (e.g. "00", "##", "??")
    };

    // Intermediate token during parsing — may carry repeat count before flattening
    struct raw_token
    {
        token_type type;
        int repeat;
        std::string lit;
        bool consumed = false; // mark as consumed by detect_fraction_seconds
        std::string exp_pattern; // exponent digit pattern for scientific tokens
    };

    // A format section — divided by ';' in the format string.
    // Up to 4 sections: [positive]; [negative]; [zero]; [text]
    struct section
    {
        std::vector<token> tokens;
        bool has_condition = false;
        enum cond_op : uint8_t { cond_none, cond_gt, cond_ge, cond_lt, cond_le, cond_eq, cond_ne };
        cond_op condition_op = cond_none;
        double condition_value = 0;
        std::string color;              // [ColorName] bracket
        int section_type = 0;           // 0=pos, 1=neg, 2=zero, 3=text
        int scale = 1;                  // divisor from trailing commas (1, 1000, 1000000, ...)
    };

    // Counts of digit placeholders (0, #, ?) on each side of the decimal point.
    // int_pattern records the L-to-R order of integer placeholder types ('0','#','?')
    // for correct zero-padding and suppression (e.g., "0#" differs from "#0").
    struct digit_counts
    {
        int int_zeros = 0, int_hashes = 0, int_qmarks = 0;
        int frac_zeros = 0, frac_hashes = 0, frac_qmarks = 0;
        int thousands = 0;
        std::string int_pattern;   // e.g. "##0", "0#?", "??#"
        std::string frac_pattern;  // e.g. "00", "0#?", "0?"
        int total_frac() const noexcept { return frac_zeros + frac_hashes + frac_qmarks; }
        int total_int() const noexcept { return int_zeros + int_hashes + int_qmarks; }
    };

    // Fraction format layout: positions of digit groups around the '/' separator
    struct fraction_layout
    {
        int slash_pos = -1;     // position of '/' token
        int space_pos = -1;     // position of space separator (indicates integer part exists)
        digit_counts int_counts;  // detailed digit counts for integer part (0/#/? + thousands)
        int num_digits = 0;     // count of numerator-placeholder tokens
        int den_digits = 0;     // count of denominator-placeholder tokens
        int num_zeros = 0;      // count of 0 tokens in numerator (for zero-padding)
        int num_qmarks = 0;     // count of ? tokens in numerator (for space padding)
        int den_zeros = 0;      // count of 0 tokens in denominator (for zero-padding)
        int den_qmarks = 0;     // count of ? tokens in denominator (for space padding)
    };

    // Summary of what a section contains (for dispatching to the right formatter)
    // Fraction layout is computed eagerly to avoid a second token scan in the formatter.
    struct section_info
    {
        bool has_digits = false;
        bool has_date_time = false;
        int percent_count = 0;
        bool has_scientific = false;
        bool is_fraction = false;
        bool has_ampm = false;
        fraction_layout frac_layout;
    };

    // ---- data ----

    std::vector<section> sections;
    bool is_general = false;       // Format is "General" (no formatting)
    int text_section_idx = -1;     // Index of the 4th (text) section, or -1

    // =========================================================================
    // Token type predicates
    // =========================================================================

    static bool is_digit_token(token_type t) noexcept
    {
        return t == token_type::digit_zero || t == token_type::digit_hash || t == token_type::digit_qmark;
    }

    static bool is_date_token(token_type t) noexcept
    {
        switch (t)
        {
        case token_type::year_2: case token_type::year_4:
        case token_type::month_n: case token_type::month_nn:
        case token_type::month_mmm: case token_type::month_mmmm:
        case token_type::month_mmmmm:
        case token_type::day_d: case token_type::day_dd:
        case token_type::day_ddd: case token_type::day_dddd:
            return true;
        default: return false;
        }
    }

    static bool is_time_token(token_type t) noexcept
    {
        switch (t)
        {
        case token_type::hour_h: case token_type::hour_hh:
        case token_type::minute_m: case token_type::minute_mm:
        case token_type::second_s: case token_type::second_ss:
        case token_type::am_pm:
        case token_type::elapsed_hours: case token_type::elapsed_minutes:
        case token_type::elapsed_seconds:
        case token_type::frac_second:
            return true;
        default: return false;
        }
    }

    // =========================================================================
    // Parse — ECMA-376 §18.8.31 format code grammar
    //
    // Format string structure:
    //   [condition] [color] [$-locale] format-body
    //
    // Format body token types:
    //   "text"   — literal (quoted)
    //   \c       — escaped character
    //   _c       — skip width of character c
    //   *c       — fill with character c
    //   @        — text placeholder
    //   0 # ?    — digit placeholders
    //   .        — decimal point
    //   ,        — thousands separator (or scale if after last digit)
    //   %        — percent
    //   E+ E-    — scientific notation
    //   y m d h s — date/time components
    //   AM/PM A/P — 12-hour clock indicator
    //   [h] [m] [s] [hh] [mm] [ss] — elapsed time
    //   /        — fraction separator
    //   :        — time separator
    //   () - + space — literal characters
    // =========================================================================

    void parse(const std::string& fmt)
    {
        sections.clear();
        is_general = false;
        text_section_idx = -1;

        if (fmt.empty()) { is_general = true; return; }

        // ECMA-376: "General" (case-insensitive) means no formatting.
        // Strip leading bracket expressions (color, condition, locale) before checking.
        // E.g., "General", "[Red]General", "[$USD-409]General", "[>10]General" all → General.
        {
            size_t pos = 0;
            while (pos < fmt.size() && fmt[pos] == '[')
            {
                const size_t end = fmt.find(']', pos);
                if (end == std::string::npos) break;
                pos = end + 1;
            }
            const std::string body = fmt.substr(pos);
            if (body.size() == 7)
            {
                bool match = true;
                for (size_t i = 0; i < 7; ++i)
                    if (static_cast<char>(std::toupper(static_cast<unsigned char>(body[i]))) != "GENERAL"[i])
                    { match = false; break; }
                if (match) { is_general = true; return; }
            }
        }

        auto raw_sections = split_sections(fmt);
        if (raw_sections.size() > 4)
            throw std::runtime_error("too many format sections (max 4)");

        int condition_count = 0;
        for (size_t i = 0; i < raw_sections.size(); ++i)
        {
            auto sec = parse_section(raw_sections[i], static_cast<int>(i), condition_count);
            sections.push_back(std::move(sec));
        }

        if (sections.empty())
        {
            section s;
            s.tokens.push_back({token_type::literal, "", 0});
            sections.push_back(std::move(s));
        }

        // ECMA-376: 4th section is the text section (used when formatting strings)
        text_section_idx = (sections.size() >= 4) ? 3 : -1;
    }

private:
    // Split the format string by ';' into sections.
    // Respects: quoted strings, bracket groups, and escape sequences.
    static std::vector<std::string> split_sections(const std::string& fmt)
    {
        std::vector<std::string> result;
        std::string cur;
        bool in_quote = false, in_bracket = false;

        for (size_t i = 0; i < fmt.size(); ++i)
        {
            const char c = fmt[i];
            if (c == '"' && !in_bracket)
            {
                cur += c;
                if (in_quote && i + 1 < fmt.size() && fmt[i + 1] == '"')
                {
                    // ECMA-376: "" within a quoted string is an escaped quote
                    cur += fmt[++i];
                    continue;
                }
                in_quote = !in_quote;
                continue;
            }
            if (!in_quote && c == '[')         { in_bracket = true; cur += c; continue; }
            if (!in_quote && c == ']')         { in_bracket = false; cur += c; continue; }
            if (!in_quote && !in_bracket && c == ';') { result.push_back(cur); cur.clear(); continue; }
            if (c == '\\' && i + 1 < fmt.size())      { cur += c; cur += fmt[++i]; continue; }
            cur += c;
        }
        result.push_back(cur);
        return result;
    }

    // Parse one section from its raw string (after bracket removal).
    // Pipeline: bracket → tokenize → resolve m/mm → detect frac seconds → flatten → scale
    section parse_section(const std::string& raw, int section_index, int& condition_count)
    {
        section sec;
        sec.section_type = section_index;
        size_t pos = 0;

        // Consume all prefix brackets: conditions, colors, locales.
        // Elapsed time brackets are NOT consumed here — they're inline tokens
        // handled by tokenize_section below.
        while (pos < raw.size() && raw[pos] == '[')
        {
            const size_t end = raw.find(']', pos);
            if (end == std::string::npos) break;
            const std::string inner = raw.substr(pos + 1, end - pos - 1);
            if (is_elapsed_bracket(inner) != '\0') break;
            parse_bracket(raw, pos, sec, condition_count);
        }
        auto raw_tokens = tokenize_section(raw, pos);
        resolve_mm_ambiguity(raw_tokens);
        detect_fraction_seconds(raw_tokens);
        flatten_tokens(raw_tokens, sec.tokens);
        suppress_extra_fill_tokens(sec.tokens);
        sec.scale = compute_scale(sec.tokens);
        return sec;
    }

    // Parse a leading [bracket] expression.
    // Three forms:
    //   [condition]    — e.g., [>100], [<0], [=5], [>=50]
    //   [ColorName]    — e.g., [Red], [Blue], [Color 3]
    //   [$-locale]     — e.g., [$USD-409], [$¥-804]
    // Elapsed time brackets [h]/[hh]/[m]/[mm]/[s]/[ss] are filtered out by
    // parse_section before this function is called.
    static void parse_bracket(const std::string& raw, size_t& pos, section& sec, int& condition_count)
    {
        if (pos >= raw.size() || raw[pos] != '[') return;

        const size_t end = raw.find(']', pos);
        if (end == std::string::npos) return;

        const std::string bracket = raw.substr(pos + 1, end - pos - 1);
        pos = end + 1;

        const bool is_locale = (!bracket.empty() && bracket[0] == '$');

        if (!is_locale)
        {
            // Condition: starts with >, <, or =
            if (!bracket.empty() && (bracket[0] == '>' || bracket[0] == '<' || bracket[0] == '='))
            {
                if (sec.has_condition)
                    throw std::runtime_error("multiple conditions in one section");
                if (++condition_count > 2)
                    throw std::runtime_error("format should have a maximum of two sections with conditions");

                sec.has_condition = true;
                const char* p = bracket.c_str();
                     if (p[0] == '>' && p[1] == '=') { sec.condition_op = section::cond_ge; p += 2; }
                else if (p[0] == '<' && p[1] == '=') { sec.condition_op = section::cond_le; p += 2; }
                else if (p[0] == '<' && p[1] == '>') { sec.condition_op = section::cond_ne; p += 2; }
                else if (p[0] == '>')                  { sec.condition_op = section::cond_gt; p += 1; }
                else if (p[0] == '<')                  { sec.condition_op = section::cond_lt; p += 1; }
                else if (p[0] == '=')                  { sec.condition_op = section::cond_eq; p += 1; }
                sec.condition_value = std::strtod(p, nullptr);
            }
            else
            {
                sec.color = bracket; // [Red], [Blue], [Color N], etc.
            }
        }
        // is_locale ([$xxx-xxx]): consumed but ignored — locale formatting not supported
    }

    // Tokenize the body of a section (after bracket removal) into raw_tokens.
    // Each raw_token may carry a repeat count for digit placeholders.
    static std::vector<raw_token> tokenize_section(const std::string& raw, size_t pos)
    {
        std::vector<raw_token> tokens;

        while (pos < raw.size())
        {
            const char c = raw[pos];

            // ---- ECMA-376 special characters ----

            if (c == '"')  // Quoted literal string
            {
                // ECMA-376: "" within a quoted string is an escaped quote.
                // Walk character-by-character to handle embedded quotes.
                std::string literal;
                ++pos; // skip opening quote
                while (pos < raw.size())
                {
                    if (raw[pos] == '"')
                    {
                        if (pos + 1 < raw.size() && raw[pos + 1] == '"')
                        {
                            // Escaped quote: "" → "
                            literal += '"';
                            pos += 2;
                        }
                        else
                        {
                            // Closing quote
                            ++pos;
                            break;
                        }
                    }
                    else
                    {
                        literal += raw[pos];
                        ++pos;
                    }
                }
                tokens.push_back({token_type::literal, 0, literal});
                continue;
            }
            if (c == '\\' && pos + 1 < raw.size())  // Escaped character
            {
                tokens.push_back({token_type::literal, 0, std::string(1, raw[pos + 1])});
                pos += 2; continue;
            }
            if (c == '_' && pos + 1 < raw.size())  // Skip width of next char
            {
                tokens.push_back({token_type::skip, 0, std::string(1, raw[pos + 1])});
                pos += 2; continue;
            }
            if (c == '*' && pos + 1 < raw.size())  // Fill with next char
            {
                tokens.push_back({token_type::fill, 0, std::string(1, raw[pos + 1])});
                pos += 2; continue;
            }
            if (c == '@')  // Text placeholder
            {
                tokens.push_back({token_type::text_placeholder, 0, ""});
                ++pos; continue;
            }

            // ---- Elapsed time: [h], [hh], [m], [mm], [s], [ss] (variable-length) ----
            // Other bracket expressions ([Color], [DBNumX], etc.) are skipped.
            if (pos < raw.size() && raw[pos] == '[')
            {
                const size_t closing = raw.find(']', pos);
                if (closing != std::string::npos && closing > pos + 1)
                {
                    const std::string inner = raw.substr(pos + 1, closing - pos - 1);
                    const char et = is_elapsed_bracket(inner);
                    if (et != '\0')
                    {
                        const int n = static_cast<int>(inner.size());
                        if (et == 'h') { tokens.push_back({token_type::elapsed_hours, n, ""});   pos = closing + 1; continue; }
                        if (et == 'm') { tokens.push_back({token_type::elapsed_minutes, n, ""}); pos = closing + 1; continue; }
                        if (et == 's') { tokens.push_back({token_type::elapsed_seconds, n, ""}); pos = closing + 1; continue; }
                    }
                    // Not an elapsed time bracket — skip it (e.g., [Red], [DBNum1])
                    pos = closing + 1;
                    continue;
                }
            }

            // ---- Scientific notation: E+/E-/e+/e- followed by digit placeholders (ECMA-376: E or e) ----
            if ((c == 'E' || c == 'e') && pos + 1 < raw.size() &&
                (raw[pos + 1] == '+' || raw[pos + 1] == '-'))
            {
                const bool plus = (raw[pos + 1] == '+');
                const bool upper = (c == 'E');
                pos += 2;
                std::string pattern;
                while (pos < raw.size() && (raw[pos] == '0' || raw[pos] == '#' || raw[pos] == '?'))
                { pattern += raw[pos]; ++pos; }
                if (pattern.empty()) pattern = "0";
                const int ed = static_cast<int>(pattern.size());
                auto rt = raw_token{token_type::scientific, ed,
                    std::string(1, upper ? 'E' : 'e') + (plus ? "+" : "-")};
                rt.exp_pattern = std::move(pattern);
                tokens.push_back(std::move(rt));
                continue;
            }

            // ---- Date/time components (repeat counts determine type) ----
            if (c == 'y' || c == 'Y')  // Year: yy → 2-digit, yyy+ → 4-digit
            {
                int n = 1;
                while (pos + n < raw.size() && (raw[pos + n] == 'y' || raw[pos + n] == 'Y')) ++n;
                tokens.push_back({n >= 3 ? token_type::year_4 : token_type::year_2, n, ""});
                pos += n; continue;
            }
            if (c == 'm' || c == 'M')  // Month (or Minute, resolved later): m, mm, mmm, mmmm, mmmmm+
            {
                int n = 1;
                while (pos + n < raw.size() && (raw[pos + n] == 'm' || raw[pos + n] == 'M')) ++n;
                token_type t;
                if (n >= 5)      t = token_type::month_mmmmm;
                else if (n == 4) t = token_type::month_mmmm;
                else if (n == 3) t = token_type::month_mmm;
                else             t = (n == 2) ? token_type::month_nn : token_type::month_n;
                tokens.push_back({t, n, ""});
                pos += n; continue;
            }
            if (c == 'd' || c == 'D')  // Day: d, dd, ddd, dddd+
            {
                int n = 1;
                while (pos + n < raw.size() && (raw[pos + n] == 'd' || raw[pos + n] == 'D')) ++n;
                token_type t;
                if (n >= 4)      t = token_type::day_dddd;
                else if (n == 3) t = token_type::day_ddd;
                else if (n == 2) t = token_type::day_dd;
                else             t = token_type::day_d;
                tokens.push_back({t, n, ""});
                pos += n; continue;
            }
            if (c == 'h' || c == 'H')  // Hour: h, hh
            {
                int n = 1;
                while (pos + n < raw.size() && (raw[pos + n] == 'h' || raw[pos + n] == 'H')) ++n;
                tokens.push_back({n >= 2 ? token_type::hour_hh : token_type::hour_h, n, ""});
                pos += n; continue;
            }
            if (c == 's' || c == 'S')  // Second: s, ss
            {
                int n = 1;
                while (pos + n < raw.size() && (raw[pos + n] == 's' || raw[pos + n] == 'S')) ++n;
                tokens.push_back({n >= 2 ? token_type::second_ss : token_type::second_s, n, ""});
                pos += n; continue;
            }

            // ---- AM/PM (5 chars) and A/P (3 chars) — case-insensitive per ECMA-376 ----
            if (pos + 4 < raw.size())
            {
                const std::string ampm = raw.substr(pos, 5);
                if ((ampm[0] == 'A' || ampm[0] == 'a') &&
                    (ampm[1] == 'M' || ampm[1] == 'm') &&
                    ampm[2] == '/' &&
                    (ampm[3] == 'P' || ampm[3] == 'p') &&
                    (ampm[4] == 'M' || ampm[4] == 'm'))
                { tokens.push_back({token_type::am_pm, 0, ampm}); pos += 5; continue; }
            }
            if (pos + 2 < raw.size())
            {
                const std::string ap = raw.substr(pos, 3);
                if ((ap[0] == 'A' || ap[0] == 'a') &&
                    ap[1] == '/' &&
                    (ap[2] == 'P' || ap[2] == 'p'))
                { tokens.push_back({token_type::am_pm, 0, ap}); pos += 3; continue; }
            }

            // ---- Single-character tokens ----
            if (c == '.')      { tokens.push_back({token_type::decimal, 0, ""}); ++pos; continue; }
            if (c == '%')      { tokens.push_back({token_type::percent, 0, ""}); ++pos; continue; }
            if (c == ',')      { tokens.push_back({token_type::thousands, 0, ""}); ++pos; continue; }
            if (c == ':')      { tokens.push_back({token_type::literal, 0, ":"}); ++pos; continue; }
            if (c == '/')      { tokens.push_back({token_type::literal, 0, "/"}); ++pos; continue; }

            // ---- Digit placeholders (with repeat counts) ----
            if (c == '0')
            {
                int n = 1;
                while (pos + n < raw.size() && raw[pos + n] == '0') ++n;
                tokens.push_back({token_type::digit_zero, n, ""});
                pos += n; continue;
            }
            if (c >= '1' && c <= '9')
            {
                tokens.push_back({token_type::literal, 0, std::string(1, c)});
                ++pos; continue;
            }
            if (c == '#')
            {
                int n = 1;
                while (pos + n < raw.size() && raw[pos + n] == '#') ++n;
                tokens.push_back({token_type::digit_hash, n, ""});
                pos += n; continue;
            }
            if (c == '?')
            {
                int n = 1;
                while (pos + n < raw.size() && raw[pos + n] == '?') ++n;
                tokens.push_back({token_type::digit_qmark, n, ""});
                pos += n; continue;
            }

            // ---- Characters that are literal without escaping (§18.8.31) ----
            if (c == '$' || c == '(' || c == ')' || c == '-' || c == '+' || c == ' ' ||
                c == '!' || c == '^' || c == '&' || c == '\'' || c == '~' ||
                c == '{' || c == '}' || c == '<' || c == '>' || c == '=' ||
                c == '`' || c == '|')
            {
                tokens.push_back({token_type::literal, 0, std::string(1, c)});
                ++pos; continue;
            }

            // Any other character not explicitly handled above is treated as a
            // literal — ECMA-376 defines special characters, everything else
            // displays as entered.
            tokens.push_back({token_type::literal, 0, std::string(1, c)});
            ++pos;
        }
        return tokens;
    }

    // Resolve the m/mm ambiguity: are they months or minutes?
    //
    // ECMA-376 §18.8.31: m/mm immediately after h/hh or immediately before
    // s/ss is a minute.  The standard examples use ":" as the time separator
    // (e.g. h:mm:ss, [h]:mm:ss), so m/mm adjacent to ":" is also resolved as
    // minutes — this is the only logical interpretation consistent with the
    // standard's own examples.
    static void resolve_mm_ambiguity(std::vector<raw_token>& tokens)
    {
        for (size_t i = 0; i < tokens.size(); ++i)
        {
            auto& rt = tokens[i];
            if (rt.type != token_type::month_n && rt.type != token_type::month_nn) continue;

            const bool after_h  = (i > 0 && (tokens[i - 1].type == token_type::hour_h || tokens[i - 1].type == token_type::hour_hh || tokens[i - 1].type == token_type::elapsed_hours));
            const bool before_s = (i + 1 < tokens.size() &&
                (tokens[i + 1].type == token_type::second_s || tokens[i + 1].type == token_type::second_ss));
            const bool col_before = (i > 0 && tokens[i - 1].type == token_type::literal && tokens[i - 1].lit == ":");
            const bool col_after  = (i + 1 < tokens.size() && tokens[i + 1].type == token_type::literal && tokens[i + 1].lit == ":");

            if (after_h || before_s || col_before || col_after)
                rt.type = (rt.type == token_type::month_nn) ? token_type::minute_mm : token_type::minute_m;
        }
    }

    // Detect fractional seconds: ".000" immediately after "s" or "ss".
    // Converts the decimal token to frac_second and marks the following zeros
    // as consumed so flatten_tokens will skip them.
    static void detect_fraction_seconds(std::vector<raw_token>& tokens)
    {
        for (size_t i = 0; i < tokens.size(); ++i)
        {
            if (tokens[i].type != token_type::decimal) continue;
            if (i == 0) continue;
            const auto prev = tokens[i - 1].type;
            if (prev != token_type::second_s && prev != token_type::second_ss
                && prev != token_type::elapsed_seconds
                && prev != token_type::elapsed_hours
                && prev != token_type::elapsed_minutes) continue;

            int nz = 0;
            size_t j = i + 1;
            while (j < tokens.size() && tokens[j].type == token_type::digit_zero)
            { nz += tokens[j].repeat; ++j; }

            if (nz > 0)
            {
                tokens[i].type = token_type::frac_second;
                tokens[i].repeat = nz;
                // Mark consumed zero tokens so flatten_tokens skips them
                for (size_t k = i + 1; k < j; ++k)
                    tokens[k].consumed = true;
            }
        }
    }

    // Expand raw_tokens with repeat counts into individual tokens.
    // E.g., "000" (digit_zero, repeat=3) → 3 individual digit_zero tokens.
    //
    // Tokens whose repeat count encodes zero-padding width propagate it to the
    // flattened token. These are all date/time tokens where the number of
    // letters determines the minimum display width:
    //   h/hh, m/mm, s/ss  →  repeat = zero-padded width
    //   yyyy, yyyyy, …     →  year (4+ digits, zero-padded to repeat)
    //   .000 after s        →  fractional seconds (width = repeat)
    //   [hh], [mm], [ss]   →  elapsed time (width = repeat)
    //
    // NOT propagated (repeat indicates type variant, not padding width):
    //   yy (fixed 2-digit), d/dd/ddd/dddd (day number vs. name),
    //   mmm/mmmm/mmmmm (month name abbreviation level)
    static bool is_repeat_propagated(token_type tt) noexcept
    {
        switch (tt)
        {
        case token_type::frac_second:
        case token_type::elapsed_hours: case token_type::elapsed_minutes:
        case token_type::elapsed_seconds:
        case token_type::year_4:
        case token_type::hour_h:   case token_type::hour_hh:
        case token_type::minute_m: case token_type::minute_mm:
        case token_type::second_s: case token_type::second_ss:
            return true;
        default: return false;
        }
    }

    static void flatten_tokens(const std::vector<raw_token>& raw_tokens, std::vector<token>& out)
    {
        for (auto& rt : raw_tokens)
        {
            if (rt.consumed) continue;
            if (rt.type == token_type::literal && rt.lit.empty()) continue;

            const int count = is_digit_token(rt.type) ? rt.repeat : 1;
            for (int r = 0; r < count; ++r)
            {
                token t;
                t.type = rt.type;
                t.repeat = is_repeat_propagated(rt.type) ? rt.repeat : 1;

                if (rt.type == token_type::literal || rt.type == token_type::skip ||
                    rt.type == token_type::fill || rt.type == token_type::am_pm ||
                    rt.type == token_type::text_placeholder ||
                    is_digit_token(rt.type))
                    t.literal = rt.lit;
                if (rt.type == token_type::scientific)
                {
                    t.exp_plus_sign = (rt.lit[1] == '+');
                    t.exp_upper_case = (rt.lit[0] == 'E');
                    t.exp_digits = rt.repeat;
                    t.exp_pattern = rt.exp_pattern;
                }
                out.push_back(t);
            }
        }
    }

    // ECMA-376: "If more than one asterisk appears in one section of the
    // format, all but the last asterisk shall be ignored."
    // Remove all fill tokens except the last one.
    static void suppress_extra_fill_tokens(std::vector<token>& tokens) noexcept
    {
        int last_fill = -1;
        for (size_t i = 0; i < tokens.size(); ++i)
            if (tokens[i].type == token_type::fill) last_fill = static_cast<int>(i);

        if (last_fill < 0) return;

        size_t write = 0;
        for (size_t i = 0; i < tokens.size(); ++i)
        {
            if (tokens[i].type == token_type::fill && static_cast<int>(i) != last_fill)
                continue; // ignored per ECMA-376
            if (write != i)
                tokens[write] = std::move(tokens[i]);
            ++write;
        }
        tokens.resize(write);
    }

    // Compute the scale divisor from commas.
    // ECMA-376: "If the number format contains a comma (,) to the left of the
    // digit placeholder, then the number shall be divided by 1000 for each
    // comma."  This covers three cases:
    //
    //   1. Leading commas  — commas before the first digit placeholder
    //   2. Trailing commas — commas after the last digit placeholder
    //   3. Both (e.g., ",#,##0,") — product of leading + trailing
    //
    // E.g., "#,##0,," → scale = 1,000,000   (trailing)
    //        ",0"     → scale = 1,000        (leading)
    //        ",0,"    → scale = 1,000,000    (leading + trailing)
    //        ",#,##0" → scale = 1,000        (leading; middle comma is grouping)
    //        "#,##0,,.00" → scale = 1,000,000
    //        "0.0,"   → scale = 1,000        (trailing on frac side)
    //
    // Only ECMA-376 digit placeholders (0, #, ?) qualify as the "digit".
    // Literal characters (including literal digits 1-9) do not reset the anchor.
    static int compute_scale(const std::vector<token>& tokens)
    {
        // Find the decimal point position (or end of tokens)
        int decimal_pos = static_cast<int>(tokens.size());
        for (int i = 0; i < static_cast<int>(tokens.size()); ++i)
            if (tokens[i].type == token_type::decimal) { decimal_pos = i; break; }

        // Find boundaries of digit placeholders on the integer side
        int first_int_digit = -1;
        int last_int_digit = -1;
        for (int i = 0; i < decimal_pos; ++i)
        {
            if (is_digit_token(tokens[i].type))
            {
                if (first_int_digit < 0) first_int_digit = i;
                last_int_digit = i;
            }
        }

        // Find rightmost digit placeholder on the fractional side
        int last_frac_digit = -1;
        for (int i = decimal_pos + 1; i < static_cast<int>(tokens.size()); ++i)
            if (is_digit_token(tokens[i].type))
                last_frac_digit = i;

        int scale = 1;

        // Leading commas: before the first integer digit, walking left
        if (first_int_digit >= 0)
        {
            for (int i = first_int_digit - 1; i >= 0; --i)
            {
                if (tokens[i].type == token_type::thousands) scale *= 1000;
                else break;
            }
        }

        // Trailing commas: after the last integer digit (up to decimal point)
        if (last_int_digit >= 0)
        {
            for (int i = last_int_digit + 1; i < decimal_pos; ++i)
            {
                if (tokens[i].type == token_type::thousands) scale *= 1000;
                else break;
            }
        }

        // Fractional side: commas after the last fractional digit
        if (last_frac_digit >= 0)
        {
            for (int i = last_frac_digit + 1; i < static_cast<int>(tokens.size()); ++i)
            {
                if (tokens[i].type == token_type::thousands) scale *= 1000;
                else break;
            }
        }

        return scale;
    }

public:
    // =========================================================================
    // Section selection — ECMA-376 §18.8.31 section ordering
    //
    // Priority:
    //   1. Condition sections ([>100], [<0], etc.) — first match wins
    //   2. If conditions exist but none matched → first unconditional section
    //   3. Standard 1/2/3-section logic:
    //      1 section  → applies to all numbers
    //      2 sections → [0] positive/zero, [1] negative
    //      3 sections → [0] positive, [1] negative, [2] zero
    // =========================================================================

    const section* select_section(double number) const noexcept
    {
        if (sections.empty()) return nullptr;

        // Step 1: evaluate condition sections
        bool any_condition = false;
        for (auto& sec : sections)
            if (sec.has_condition) { any_condition = true; break; }

        if (any_condition)
        {
            for (auto& sec : sections)
            {
                if (!sec.has_condition) continue;
                if (condition_matches(sec, number)) return &sec;
            }
            // No condition matched → fall back to first unconditional section
            for (auto& sec : sections)
                if (!sec.has_condition) return &sec;
            // ECMA-376: all sections have conditions but none matched
            return nullptr;
        }

        // Step 2: standard positive/negative/zero selection
        const size_t n = sections.size();
        if (n == 1)       return &sections[0];
        else if (n == 2)  return (number < 0) ? &sections[1] : &sections[0];
        else              return number > 0 ? &sections[0] : number < 0 ? &sections[1] : &sections[2];
    }

private:
    static bool condition_matches(const section& sec, double number) noexcept
    {
        // Scale epsilon relative to the condition value to handle both small
        // and large thresholds correctly (e.g., [=1e-8] vs [=1e10]).
        const double eps = float_epsilon * std::fmax(1.0, std::fabs(sec.condition_value));
        switch (sec.condition_op)
        {
        case section::cond_gt: return number > sec.condition_value + eps;
        case section::cond_ge: return number >= sec.condition_value - eps;
        case section::cond_lt: return number < sec.condition_value - eps;
        case section::cond_le: return number <= sec.condition_value + eps;
        case section::cond_eq: return std::fabs(number - sec.condition_value) < eps;
        case section::cond_ne: return std::fabs(number - sec.condition_value) >= eps;
        default: return false;
        }
    }

public:
    // =========================================================================
    // Section analysis — classify what kind of formatting a section contains
    // =========================================================================

    static section_info analyze_section(const section& sec) noexcept
    {
        section_info info;
        for (size_t i = 0; i < sec.tokens.size(); ++i)
        {
            const auto& tok = sec.tokens[i];
            const auto tt = tok.type;

            if (is_digit_token(tt) || is_frac_literal_digit(tok))
            {
                info.has_digits = true;
                if (!info.is_fraction)
                    info.is_fraction = is_fraction_denominator(sec.tokens, i);
            }
            else if (tt == token_type::percent)
                info.percent_count++;
            else if (tt == token_type::scientific)
                info.has_scientific = true;
            else if (tt == token_type::am_pm)
                info.has_ampm = true;
            else if (is_date_token(tt) || is_time_token(tt))
                info.has_date_time = true;
        }
        // Compute fraction layout eagerly to avoid a redundant full token scan
        if (info.is_fraction)
            info.frac_layout = analyze_fraction_layout(sec);
        return info;
    }

private:
    // A digit token is a fraction denominator if it follows "/" and
    // there is a digit before the "/" (the numerator).
    static bool is_fraction_denominator(const std::vector<token>& tokens, size_t pos) noexcept
    {
        if (pos < 2) return false;
        if (tokens[pos - 1].type != token_type::literal || tokens[pos - 1].literal != "/")
            return false;
        return is_digit_token(tokens[pos - 2].type) || is_frac_literal_digit(tokens[pos - 2]);
    }

    // Literal digits 1-9 are fraction placeholders (e.g., "# ?/8")
    static bool is_frac_literal_digit(const token& tok) noexcept
    {
        return tok.type == token_type::literal && tok.literal.size() == 1
            && tok.literal[0] >= '1' && tok.literal[0] <= '9';
    }

    // Append common tokens (literal, skip, fill) to the output.
    // Returns true if the token was handled, false if the caller should process it.
    static bool append_common_token(std::string& out, const token& tok) noexcept
    {
        switch (tok.type)
        {
        case token_type::literal: out += tok.literal; return true;
        case token_type::skip:    out += ' '; return true;
        case token_type::fill:    out += tok.literal; return true;
        default: return false;
        }
    }

public:
    // =========================================================================
    // format: double — main entry point
    // =========================================================================

    std::string format_double(double number, bool date1904) const
    {
        // ECMA-376: NaN and INF both display as #NUM!
        if (std::isnan(number) || std::isinf(number)) return "#NUM!";

        if (is_general || sections.empty())
            return format_number_general(number);

        const section* sec = select_section(number);
        // ECMA-376: "If the cell value does not meet any of the criteria, then
        // pound signs ("#") are displayed across the width of the cell."
        if (!sec) return std::string(default_cell_width, '#');

        const section_info info = analyze_section(*sec);

        // No digit or date/time tokens → pure literal template (handles text sections with @)
        if (!info.has_digits && !info.has_date_time)
            return render_literal_template(*sec, number);

        // Check if the selected section provides its own sign for negative numbers.
        // ECMA-376: If the section begins with "-" or "(" (possibly preceded by
        // skip/fill tokens like "_(*"), the section owns the sign and no automatic
        // minus is prepended.  This applies to the actually selected section for
        // this negative value, regardless of its index (important when conditional
        // sections reorder the effective negative section).
        bool neg_section_owns_sign = false;
        if (number < 0)
        {
            for (auto& tok : sec->tokens)
            {
                if (tok.type == token_type::skip || tok.type == token_type::fill)
                    continue;
                if (tok.type == token_type::literal)
                {
                    if (tok.literal == "-" || tok.literal == "(")
                    { neg_section_owns_sign = true; break; }
                }
                break; // first non-skip non-fill token is not a sign → no sign owned
            }
        }

        // ECMA-376: negative dates/times are not valid — display as "###########"
        if (number < 0 && info.has_date_time)
            return std::string(default_cell_width, '#');

        // Dispatch to the appropriate specialized formatter
        if (info.has_date_time)  return format_date_time_section(*sec, std::fabs(number), date1904, info.has_ampm);
        if (info.is_fraction)    return format_fraction_section(*sec, info, number, neg_section_owns_sign, info.percent_count);
        if (info.has_scientific) return format_scientific_section(*sec, number, neg_section_owns_sign, info.percent_count);

        return format_regular_number_section(*sec, number, neg_section_owns_sign, info.percent_count);
    }

private:
    // Render a section that has no digit or date/time tokens — just literals,
    // text placeholders, and skip tokens.
    static std::string render_literal_template(const section& sec, double number)
    {
        std::string result;
        for (auto& tok : sec.tokens)
        {
            if (tok.type == token_type::text_placeholder) result += format_number_general(number);
            else append_common_token(result, tok);
        }
        return result;
    }

public:
    // =========================================================================
    // format: string — ECMA-376 text section formatting
    //
    // Uses the 4th section if available, otherwise the 1st section.
    // If the section contains @, the text replaces @. Otherwise, if the
    // section contains only literals/skips/fills, it's used as a template.
    // If it contains digit placeholders, the raw text is returned unchanged.
    // =========================================================================

    std::string format_text(const std::string& text) const
    {
        if (is_general || sections.empty()) return text;

        const section* text_sec = (text_section_idx >= 0)
            ? &sections[static_cast<size_t>(text_section_idx)]
            : &sections[0];

        if (text_sec->tokens.empty()) return text;

        std::string result;
        bool has_placeholder = false;
        for (auto& tok : text_sec->tokens)
        {
            if (tok.type == token_type::text_placeholder) { result += text; has_placeholder = true; }
            else append_common_token(result, tok);
        }

        // If no @ placeholder, check if the section is a pure template
        if (!has_placeholder)
        {
            for (auto& tok : text_sec->tokens)
                if (!(tok.type == token_type::literal || tok.type == token_type::skip || tok.type == token_type::fill))
                    return text; // contains digit/date tokens → return raw text
        }
        return result;
    }

    // =========================================================================
    // format: date/time section
    // =========================================================================

    std::string format_date_time_section(const section& sec, double raw_value, bool date1904, bool has_ampm) const
    {
        const date_parts dp = serial_to_date(raw_value, date1904);
        const size_t month_idx = static_cast<size_t>(dp.month) - 1;

        // Fractional seconds value — may be overridden by elapsed_seconds below
        double frac_sec = dp.fsec;

        std::string result;
        result.reserve(sec.tokens.size() * 4);

        for (auto& tok : sec.tokens)
        {
            switch (tok.type)
            {
            case token_type::year_2:      append_padded(result, dp.year % 100, 2); break;
            case token_type::year_4:      append_padded(result, dp.year, tok.repeat); break;
            case token_type::month_n:     append_int(result, dp.month); break;
            case token_type::month_nn:    append_padded(result, dp.month, 2); break;
            case token_type::month_mmm:   result += month_abbr[month_idx]; break;
            case token_type::month_mmmm:  result += month_names[month_idx]; break;
            case token_type::month_mmmmm: result += month_names[month_idx][0]; break;
            case token_type::day_d:       append_int(result, dp.day); break;
            case token_type::day_dd:      append_padded(result, dp.day, 2); break;
            case token_type::day_ddd:     result += day_abbr[static_cast<size_t>(dp.dow)]; break;
            case token_type::day_dddd:    result += day_names[static_cast<size_t>(dp.dow)]; break;
            case token_type::hour_h:
            case token_type::hour_hh: {
                const int h = has_ampm ? (dp.hour % 12 == 0 ? 12 : dp.hour % 12) : dp.hour;
                append_padded(result, h, tok.repeat); break;
            }
            case token_type::minute_m:
            case token_type::minute_mm:   append_padded(result, dp.minute, tok.repeat); break;
            case token_type::second_s:
            case token_type::second_ss:
                append_padded(result, dp.second, tok.repeat);
                frac_sec = dp.fsec; // reset to fractional seconds after elapsed hours/minutes
                break;
            case token_type::am_pm:       append_ampm(result, tok.literal, dp.hour >= 12); break;
            case token_type::frac_second: append_frac_second(result, frac_sec, tok.repeat); break;
            case token_type::elapsed_hours: {
                const double total_hrs = raw_value * 24;
                append_padded(result, static_cast<long long>(std::floor(total_hrs)), tok.repeat);
                frac_sec = total_hrs - std::floor(total_hrs);
                break;
            }
            case token_type::elapsed_minutes: {
                const double total_mins = raw_value * 24 * 60;
                append_padded(result, static_cast<long long>(std::floor(total_mins)), tok.repeat);
                frac_sec = total_mins - std::floor(total_mins);
                break;
            }
            case token_type::elapsed_seconds: {
                const double total_secs = raw_value * 24 * 3600;
                append_padded(result, static_cast<long long>(std::floor(total_secs)), tok.repeat);
                frac_sec = total_secs - std::floor(total_secs);
                break;
            }
            default:
                if (!append_common_token(result, tok))
                {
                    // Output non-date tokens as their literal text equivalents
                    switch (tok.type)
                    {
                    case token_type::percent:  result += '%'; break;
                    case token_type::thousands: result += ','; break;
                    case token_type::decimal:  result += '.'; break;
                    case token_type::digit_zero:  result += '0'; break;
                    case token_type::digit_hash:  result += '#'; break;
                    case token_type::digit_qmark: result += '?'; break;
                    case token_type::text_placeholder: result += '@'; break;
                    case token_type::scientific:
                        result += tok.exp_upper_case ? 'E' : 'e';
                        result += tok.exp_plus_sign ? '+' : '-';
                        result += tok.exp_pattern;
                        break;
                    default: break;
                    }
                }
                break;
            }
        }
        return result;
    }

private:
    // Append AM/PM or A/P string, preserving the original case pattern.
    // For "AM/PM" (5-char): extract the AM or PM part preserving case.
    // For "A/P" (3-char): extract the A or P part preserving case.
    static void append_ampm(std::string& out, const std::string& pattern, bool pm) noexcept
    {
        if (pattern.size() == 5) // "AM/PM" or mixed case variants
            out += pm ? pattern.substr(3, 2) : pattern.substr(0, 2);
        else // "A/P" or mixed case variants
            out += pm ? pattern.substr(2, 1) : pattern.substr(0, 1);
    }

    // Append fractional seconds with leading zeros and rounding.
    // Clamping is used instead of carrying overflow to the seconds field because
    // format_date_time_section processes tokens linearly and seconds have already
    // been emitted. The clamp only triggers for floating-point edge cases (fsec
    // extremely close to 1.0 due to accumulated double-precision error).
    static void append_frac_second(std::string& out, double fsec, int repeat) noexcept
    {
        const double pv = pow10_safe(repeat);
        int fv = static_cast<int>(std::round(fsec * pv + float_epsilon));
        if (fv >= static_cast<int>(pv)) fv = static_cast<int>(pv) - 1;

        char buf[16];
        auto r = std::to_chars(buf, buf + sizeof(buf), fv);
        const int len = static_cast<int>(r.ptr - buf);

        out += '.';
        for (int k = len; k < repeat; ++k) out += '0';
        out.append(buf, static_cast<size_t>(len));
    }

public:
    // =========================================================================
    // format: fraction section — ECMA-376 §18.8.31 fraction formats
    //
    // Examples: "# ?/?", "# ??/??", "# ?/2", "# ?/4", "# ?/8"
    //
    // Token structure (flattened):
    //   [digit_tokens_int] [space] [digit_tokens_num] [/] [digit_tokens_den]
    //   or without integer part:
    //   [digit_tokens_num] [/] [digit_tokens_den]
    //
    // Algorithm:
    //   1. Analyze fraction layout from tokens (find '/', count digit groups)
    //   2. Extract integer part and find best rational approximation of fraction
    //   3. Walk tokens, substituting digit groups with formatted values,
    //      preserving surrounding literals, skips, and fills
    // =========================================================================

    std::string format_fraction_section(const section& sec, const section_info& info, double raw_value, bool neg_section_owns_sign, int percent_count) const
    {
        const fraction_layout& fl = info.frac_layout;

        double scaled = raw_value / sec.scale;
        const bool negative = scaled < 0;
        if (negative) scaled = -scaled;
        if (percent_count > 0) scaled *= std::pow(100.0, percent_count);

        double int_part_d = std::floor(scaled);
        // Handle -0.0: std::floor(-0.0) returns -0.0, but we need 0.0
        // to avoid std::to_chars producing "-0" for the integer part.
        if (scaled == 0.0) int_part_d = 0.0;
        const double frac_part = scaled - int_part_d;

        // Determine best fraction: fixed denominator or best approximation
        long long fixed_den = compute_fixed_denominator(sec, fl);
        long long best_num, best_den;
        if (fixed_den > 0)
        {
            best_num = static_cast<long long>(std::round(frac_part * static_cast<double>(fixed_den) + float_epsilon));
            best_den = fixed_den;
        }
        else
        {
            int max_den = static_cast<int>(pow10_safe(fl.den_digits)) - 1;
            if (max_den < 1) max_den = 99;
            best_fraction(frac_part, max_den, best_num, best_den);
        }

        if (best_num == best_den) { ++int_part_d; best_num = 0; best_den = 1; }

        // ECMA-376: when there is no integer part specification (space_pos < 0),
        // the integer part is folded into the numerator as an improper fraction.
        // E.g., "?/?" with value 1.5 → "3/2", not "1/2".
        // Also handles "?/?" with value 2.0 → "2/1".
        if (fl.space_pos < 0 && int_part_d > 0)
        {
            const long long int_ll = static_cast<long long>(int_part_d);
            best_num = int_ll * best_den + best_num;
            int_part_d = 0;
        }

        const bool no_fraction = (best_num == 0);

        // Format integer part via std::to_chars on double (avoids long long overflow)
        char int_buf[64];
        auto r = std::to_chars(int_buf, int_buf + sizeof(int_buf), int_part_d, std::chars_format::fixed, 0);
        std::string int_str = format_integer_str(std::string(int_buf, r.ptr), fl.int_counts);
        bool int_suppressed = int_str.empty();

        // When the fraction part is zero, the integer must always be shown.
        // Otherwise formats like "# ?/?" produce empty output for value 0.
        if (no_fraction && int_suppressed)
        {
            int_str = (fl.int_counts.int_qmarks > 0)
                ? std::string(fl.int_counts.int_qmarks, ' ')
                : "0";
            int_suppressed = int_str.empty();
        }

        // Suppress negative sign when the rounded value is zero.
        // ECMA-376: negative values that round to zero display as "0" (not "-0").
        const bool frac_is_zero_rounded = (int_part_d == 0.0 && no_fraction);
        const bool frac_effective_negative = negative && !frac_is_zero_rounded;

        std::string num_str, den_str;
        if (!no_fraction)
        {
            char num_buf[32], den_buf[32];
            auto rn = std::to_chars(num_buf, num_buf + sizeof(num_buf), best_num);
            auto rd = std::to_chars(den_buf, den_buf + sizeof(den_buf), best_den);
            num_str.assign(num_buf, static_cast<size_t>(rn.ptr - num_buf));
            den_str.assign(den_buf, static_cast<size_t>(rd.ptr - den_buf));

            // ECMA-376: 0 placeholders force zero-padding; ? placeholders force
            // space-padding for alignment.  0 takes precedence over ?.
            // Numerator is left-padded (aligns with fraction bar on the right).
            // Denominator is left-padded with zeros, right-padded with spaces.
            if (fl.num_zeros > 0)
            {
                if (static_cast<int>(num_str.size()) < fl.num_digits)
                    num_str.insert(0, static_cast<size_t>(fl.num_digits - static_cast<int>(num_str.size())), '0');
            }
            else if (static_cast<int>(num_str.size()) < fl.num_qmarks)
            {
                num_str.insert(0, static_cast<size_t>(fl.num_qmarks - static_cast<int>(num_str.size())), ' ');
            }

            if (fl.den_zeros > 0)
            {
                if (static_cast<int>(den_str.size()) < fl.den_digits)
                    den_str.insert(0, static_cast<size_t>(fl.den_digits - static_cast<int>(den_str.size())), '0');
            }
            else if (static_cast<int>(den_str.size()) < fl.den_qmarks)
            {
                den_str.append(static_cast<size_t>(fl.den_qmarks - static_cast<int>(den_str.size())), ' ');
            }
        }

        // Walk tokens and assemble output
        std::string result;
        result.reserve(32);
        if (frac_effective_negative && !neg_section_owns_sign) result += '-';

        bool int_output = false, num_output = false, den_output = false;

        for (size_t i = 0; i < sec.tokens.size(); ++i)
        {
            const auto& tok = sec.tokens[i];
            const auto tt = tok.type;

            if (tt == token_type::thousands) continue; // handled by format_integer_str
            if (tt == token_type::percent) { result += '%'; continue; }
            if (tt == token_type::text_placeholder) { result += '@'; continue; }
            if (tt == token_type::scientific) {
                result += tok.exp_upper_case ? 'E' : 'e';
                result += tok.exp_plus_sign ? '+' : '-';
                result += tok.exp_pattern;
                continue;
            }

            if (tt == token_type::literal)
            {
                if (tok.literal == "/")
                {
                    if (!no_fraction) result += '/';
                    continue;
                }
                // Space separator between integer and fraction parts
                if (fl.space_pos >= 0 && static_cast<int>(i) == fl.space_pos && tok.literal == " ")
                {
                    if (!no_fraction && !int_suppressed) result += ' ';
                    continue;
                }
                // Literal digits 1-9 are fraction placeholders, not output as literals
                if (is_frac_literal_digit(tok))
                {
                    append_frac_placeholder(result, i, fl, int_str, num_str, den_str,
                                            no_fraction, int_output, num_output, den_output);
                    continue;
                }
                append_common_token(result, tok);
                continue;
            }
            if (append_common_token(result, tok)) continue;

            if (is_digit_token(tt))
                append_frac_placeholder(result, i, fl, int_str, num_str, den_str,
                                        no_fraction, int_output, num_output, den_output);
        }
        return result;
    }

private:
    // Compute the fixed denominator from tokens after the '/' in a fraction format.
    // Literal digits 1-9 contribute their value; placeholder digits (0, #, ?) add
    // a factor of 10. E.g., "# ?/8" → 8, "# ?/20" → 20.
    static long long compute_fixed_denominator(const section& sec, const fraction_layout& fl) noexcept
    {
        if (fl.slash_pos < 0) return 0;
        long long den = 0;
        for (int i = fl.slash_pos + 1; i < static_cast<int>(sec.tokens.size()); ++i)
        {
            const auto& dt = sec.tokens[i];
            if (is_frac_literal_digit(dt))
                den = den * 10 + (dt.literal[0] - '0');
            else if (is_digit_token(dt.type))
                den *= 10;
            else
                break;
        }
        return den;
    }

    // Append the appropriate fraction part (integer, numerator, or denominator)
    // based on the token's position relative to the '/' and space separators.
    static void append_frac_placeholder(std::string& result, size_t idx, const fraction_layout& fl,
                                        const std::string& int_str, const std::string& num_str,
                                        const std::string& den_str, bool no_fraction,
                                        bool& int_output, bool& num_output, bool& den_output) noexcept
    {
        const int i = static_cast<int>(idx);
        if (i > fl.slash_pos)
        {
            if (!no_fraction && !den_output) { result += den_str; den_output = true; }
        }
        else if (fl.space_pos >= 0 && i < fl.space_pos)
        {
            if (!int_output) { result += int_str; int_output = true; }
        }
        else
        {
            if (!no_fraction && !num_output) { result += num_str; num_output = true; }
            // ECMA-376: when there is no integer part specification and the
            // fraction is zero (no_fraction=true), the numerator tokens display
            // the integer part (e.g., "?/?" with value 0 → "0").
            else if (no_fraction && !int_output && fl.space_pos < 0) { result += int_str; int_output = true; }
        }
    }

    // Analyze the fraction layout from a flattened token vector.
    // Finds the '/' position, counts digit groups on each side,
    // and identifies the space separator that marks the integer part.
    // Literal digits 1-9 are also counted as fraction digits (e.g., # ?/8).
    static fraction_layout analyze_fraction_layout(const section& sec) noexcept
    {
        fraction_layout fl;

        // Find the '/' position
        for (size_t i = 0; i < sec.tokens.size(); ++i)
            if (sec.tokens[i].type == token_type::literal && sec.tokens[i].literal == "/")
            { fl.slash_pos = static_cast<int>(i); break; }

        if (fl.slash_pos < 0) return fl;

        // Count denominator digits (after '/')
        for (int i = fl.slash_pos + 1; i < static_cast<int>(sec.tokens.size()); ++i)
        {
            const auto& tok = sec.tokens[i];
            if (is_digit_token(tok.type) || is_frac_literal_digit(tok))
            {
                ++fl.den_digits;
                if (tok.type == token_type::digit_zero)  ++fl.den_zeros;
                if (tok.type == token_type::digit_qmark) ++fl.den_qmarks;
            }
            else
                break;
        }

        // Count numerator digits (immediately before '/')
        for (int i = fl.slash_pos - 1; i >= 0; --i)
        {
            const auto& tok = sec.tokens[i];
            if (is_digit_token(tok.type) || is_frac_literal_digit(tok))
            {
                ++fl.num_digits;
                if (tok.type == token_type::digit_zero)  ++fl.num_zeros;
                if (tok.type == token_type::digit_qmark) ++fl.num_qmarks;
            }
            else
                break;
        }

        // Check for space separator (indicating integer part exists)
        // Space is between the integer digit group and the numerator digit group
        int sep_pos = fl.slash_pos - 1 - fl.num_digits;
        if (sep_pos >= 0 && sec.tokens[sep_pos].type == token_type::literal &&
            sec.tokens[sep_pos].literal == " ")
        {
            fl.space_pos = sep_pos;
            // Count integer part digit placeholders with full type breakdown.
            // Iterates right-to-left, so prepend to int_pattern for L-to-R order.
            //
            // Grouping vs. scaling commas: a comma is a grouping separator only
            // if it has digit placeholders on BOTH sides (strictly between the
            // leftmost and rightmost integer digits in the fraction's integer part).
            // Leading or trailing commas are scaling commas handled by compute_scale.
            //
            // First pass: find digit boundaries (in L-to-R index order)
            int leftmost_digit = sep_pos;
            int rightmost_digit = -1;
            for (int i = sep_pos - 1; i >= 0; --i)
            {
                const auto tt = sec.tokens[i].type;
                if (tt == token_type::digit_zero || tt == token_type::digit_hash ||
                    tt == token_type::digit_qmark || is_frac_literal_digit(sec.tokens[i]))
                {
                    leftmost_digit = i;
                }
                else if (tt != token_type::thousands)
                    break;
            }
            for (int i = sep_pos - 1; i >= 0; --i)
            {
                const auto tt = sec.tokens[i].type;
                if (tt == token_type::digit_zero || tt == token_type::digit_hash ||
                    tt == token_type::digit_qmark || is_frac_literal_digit(sec.tokens[i]))
                {
                    rightmost_digit = i; break;
                }
            }

            // Second pass: count placeholders and grouping commas
            for (int i = sep_pos - 1; i >= 0; --i)
            {
                const auto tt = sec.tokens[i].type;
                if (tt == token_type::digit_zero)
                    { ++fl.int_counts.int_zeros;  fl.int_counts.int_pattern.insert(0, 1, '0'); }
                else if (tt == token_type::digit_hash)
                    { ++fl.int_counts.int_hashes; fl.int_counts.int_pattern.insert(0, 1, '#'); }
                else if (tt == token_type::digit_qmark)
                    { ++fl.int_counts.int_qmarks; fl.int_counts.int_pattern.insert(0, 1, '?'); }
                else if (tt == token_type::thousands)
                {
                    // Grouping comma: must have digits on BOTH sides
                    if (leftmost_digit >= 0 && rightmost_digit >= 0 &&
                        i > leftmost_digit && i < rightmost_digit)
                        ++fl.int_counts.thousands;
                }
                else if (is_frac_literal_digit(sec.tokens[i]))
                    { /* literal digit 1-9 in fraction integer part — treat as '#' */ ++fl.int_counts.int_hashes; fl.int_counts.int_pattern.insert(0, 1, '#'); }
                else
                    break;
            }
        }

        return fl;
    }

public:
    // =========================================================================
    // format: scientific section — ECMA-376 §18.8.31 scientific notation
    //
    // Format: "0.00E+00" → mantissa with digit placeholders, then E+/- with digits
    //
    // Algorithm:
    //   1. Normalize number to 1.xxx * 10^exp
    //   2. Round mantissa to the precision specified by placeholders before E
    //   3. Format exponent with the specified number of digits
    // =========================================================================

    std::string format_scientific_section(const section& sec, double raw_value, bool neg_section_owns_sign, int percent_count) const
    {
        double value = raw_value / sec.scale;
        const bool negative = value < 0;
        if (negative) value = -value;
        if (percent_count > 0) value *= std::pow(100.0, percent_count);

        // Locate the scientific token (E+ or E-)
        int sci_pos = -1;
        token sci_tok;
        for (size_t i = 0; i < sec.tokens.size(); ++i)
        {
            if (sec.tokens[i].type == token_type::scientific) { sci_tok = sec.tokens[i]; sci_pos = static_cast<int>(i); break; }
        }

        // Count digit placeholders up to the E token for precision
        const auto dc = count_digit_placeholders(sec, sci_pos);
        const int total_int = dc.total_int();
        const int total_frac = dc.total_frac();

        // Determine mantissa normalization target.
        // ECMA-376 §18.8.31:
        //   - With decimal point → exactly 1 digit to the left → mantissa ∈ [1, 10).
        //   - Without decimal point → mantissa is an N-digit integer where N is the
        //     number of digit placeholders to the left of 'E'/'e'.  E.g., "00E+00"
        //     → mantissa is a 2-digit integer in [10, 100).
        const int int_digits = (total_frac > 0 || total_int <= 1) ? 1 : total_int;
        const double norm_range = pow10_safe(int_digits);

        // Compute mantissa and exponent with normalization target.
        int exponent = 0;
        double mantissa = value;
        if (mantissa != 0.0)
        {
            exponent = static_cast<int>(std::floor(std::log10(mantissa)));
            // Subtract (int_digits - 1) so mantissa ∈ [1, 10^int_digits)
            exponent -= (int_digits - 1);
            mantissa = mantissa / pow10_safe(exponent);
            // Correct if mantissa drifted due to fp error
            if (mantissa < 1.0 - 1e-12)           { mantissa *= 10.0; --exponent; }
            else if (mantissa >= norm_range - 1e-12 * norm_range) { mantissa /= 10.0; ++exponent; }
        }

        // Round mantissa to the specified fractional precision.
        // Epsilon added AFTER scaling, capped below 0.5 to prevent corrupting exact values.
        // Only applied when there are fractional digits to round.
        const double round_factor = pow10_safe(total_frac);
        const double m_scaled = mantissa * round_factor;
        const double m_eps = (total_frac > 0)
            ? std::fmin(0.499999999999999,
                std::fmax(float_epsilon, std::fabs(m_scaled) * std::numeric_limits<double>::epsilon()))
            : 0.0;
        mantissa = std::round(m_scaled + m_eps) / round_factor;

        // Handle rounding overflow: mantissa may have rounded up to the normalization
        // range boundary.  Divide by 10 and bump exponent (rather than collapsing
        // to 1.0) to preserve the correct number of significant digits.
        // E.g., "00E+00" with 9995: mantissa=99.95 rounds to 100 → divide to
        // 10.0, exponent+1 → "10E+03" (not "01E+04").
        if (mantissa >= norm_range - 1e-12 * norm_range)
        { mantissa /= 10.0; exponent += 1; }

        // Format mantissa integer part
        const long long m_int = static_cast<long long>(std::floor(mantissa));
        std::string m_int_str = format_integer_str(std::to_string(m_int), dc);

        // Format mantissa fractional part using format_decimal_str
        // to correctly handle mixed 0/#/? placeholders per ECMA-376.
        std::string m_frac_str;
        if (total_frac > 0)
        {
            const double m_frac = mantissa - m_int;
            // Epsilon added AFTER multiplication to avoid amplifying it by round_factor
            const long long fv = static_cast<long long>(std::round(m_frac * round_factor + float_epsilon));
            m_frac_str = format_decimal_str(fv, total_frac, dc);
        }

        // If all placeholders are # (no 0) and the mantissa is zero, ensure at least "0" is shown.
        if (m_int_str.empty() && m_frac_str.empty())
            m_int_str = "0";

        // Suppress negative sign when the rounded value is zero.
        // ECMA-376: negative values that round to zero display as "0E+00" (not "-0E+00").
        const bool sci_is_zero_rounded = (m_int == 0 && m_int_str == "0"
            && (m_frac_str.empty() || m_frac_str.find_first_not_of("0 ") == std::string::npos));
        const bool sci_effective_negative = negative && !sci_is_zero_rounded;

        // Format exponent using the digit pattern (0/#/?) per ECMA-376.
        // 0 → always show digit (zero-padded), # → suppress leading zeros,
        // ? → suppress leading zeros but pad with spaces.
        const int abs_exp = std::abs(exponent);
        std::string exp_str;
        if (!sci_tok.exp_pattern.empty())
        {
            // Pad to pattern length with leading zeros, then walk pattern L-to-R.
            std::string padded = std::to_string(abs_exp);
            const int pat_sz = static_cast<int>(sci_tok.exp_pattern.size());
            if (static_cast<int>(padded.size()) < pat_sz)
                padded.insert(0, static_cast<size_t>(pat_sz - static_cast<int>(padded.size())), '0');

            // Find first non-zero digit
            size_t first_nz = 0;
            while (first_nz < padded.size() && padded[first_nz] == '0') ++first_nz;

            exp_str.reserve(padded.size());
            for (size_t i = 0; i < padded.size(); ++i)
            {
                // ECMA-376: when exponent has more digits than the pattern,
                // treat extra positions as '0' placeholders (always show).
                const char pat = (i < sci_tok.exp_pattern.size()) ? sci_tok.exp_pattern[i] : '0';
                if (i < first_nz)
                {
                    if (pat == '#') continue;       // suppress
                    if (pat == '?') exp_str += ' ';  // space for alignment
                    else exp_str += padded[i];        // '0' placeholder
                }
                else
                {
                    exp_str += padded[i];
                }
            }
        }
        else
        {
            exp_str = std::to_string(abs_exp);
            if (static_cast<int>(exp_str.size()) < sci_tok.exp_digits)
                exp_str.insert(0, static_cast<size_t>(sci_tok.exp_digits - static_cast<int>(exp_str.size())), '0');
        }
        // ECMA-376: when the exponent pattern is all '#' and abs_exp == 0,
        // all digits would be suppressed.  Ensure at least one digit remains.
        if (exp_str.empty()) exp_str = "0";

        // Walk tokens to build output, preserving surrounding text (literals, skips, fills)
        std::string result;
        result.reserve(m_int_str.size() + m_frac_str.size() + exp_str.size() + 16);
        if (sci_effective_negative && !neg_section_owns_sign) result += '-';

        bool int_output = false, past_decimal = false;
        int frac_pos = 0;

        for (size_t i = 0; i < sec.tokens.size(); ++i)
        {
            const auto& tok = sec.tokens[i];
            const auto tt = tok.type;

            if (append_common_token(result, tok)) continue;
            if (tt == token_type::percent)  { result += '%'; continue; }
            if (tt == token_type::thousands)  continue; // handled by format_integer_str
            if (tt == token_type::text_placeholder) { result += '@'; continue; }

            if (tt == token_type::decimal && static_cast<int>(i) < sci_pos)
            {
                result += '.';
                past_decimal = true;
                continue;
            }

            if (tt == token_type::scientific)
            {
                result += sci_tok.exp_upper_case ? 'E' : 'e';
                if (exponent >= 0) { if (sci_tok.exp_plus_sign) result += '+'; }
                else result += '-';
                result += exp_str;
                continue;
            }

            if (is_digit_token(tt))
            {
                if (static_cast<int>(i) < sci_pos)
                {
                    if (!past_decimal && !int_output) { result += m_int_str; int_output = true; }
                    else if (past_decimal)
                    {
                        if (frac_pos < static_cast<int>(m_frac_str.size())) { result += m_frac_str[frac_pos]; ++frac_pos; }
                        else if (tt == token_type::digit_qmark) result += ' ';
                    }
                }
            }
        }
        // ECMA-376: when there are no integer digit placeholders but the mantissa
        // integer part is non-empty, prepend it before the decimal point.
        // E.g., "$".##E+00 with value 1.5 → "$1.5E+00", not "1$.5E+00".
        if (!int_output && !m_int_str.empty())
        {
            const size_t dot = result.find('.');
            const size_t ins = (dot != std::string::npos) ? dot : 0;
            result.insert(ins, m_int_str);
        }
        return result;
    }

    // =========================================================================
    // format: regular number section — standard numeric formatting
    //
    // Handles: digit placeholders (0, #, ?), decimal point, thousands separator,
    // percent sign, and scale from trailing commas.
    //
    // Algorithm:
    //   1. Scale and round the number to the specified precision
    //   2. Format integer part (leading zeros, ? → space padding, thousands separators)
    //   3. Format fraction part (trailing zero trimming, ? → space padding)
    //   4. Walk tokens and interpolate the formatted parts
    // =========================================================================

    std::string format_regular_number_section(const section& sec, double raw_value, bool neg_section_owns_sign, int percent_count) const
    {
        // Scale and handle sign
        double value = raw_value / sec.scale;
        const bool negative = value < 0;
        if (negative) value = -value;
        if (percent_count > 0) value *= std::pow(100.0, percent_count);

        // Analyze digit placeholders
        const auto dc = count_digit_placeholders(sec);
        const int total_frac = dc.total_frac();

        // Round to the specified precision.
        //
        // A tiny epsilon is added AFTER scaling to compensate for binary
        // floating-point representation errors (e.g., 9.995 stored as
        // ~9.994999... produces 999.4999... instead of 999.5).
        //
        // Adding before:  (x + eps) * k  distorts near-half-integer values.
        // Adding after:   x * k + eps   preserves the exact threshold.
        //
        // The epsilon scales with the ULP of the scaled value (2^-52 ≈ 2.22e-16)
        // so it remains effective for all value ranges, but is capped just
        // below 0.5 so it can never push a value across the rounding threshold.
        // Only applied when there are fractional digits to round (total_frac > 0).
        const double round_factor = pow10_safe(total_frac);
        const double scaled = value * round_factor;
        const double eps = (total_frac > 0)
            ? std::fmin(0.499999999999999,
                std::fmax(float_epsilon, std::fabs(scaled) * std::numeric_limits<double>::epsilon()))
            : 0.0;
        value = std::round(scaled + eps) / round_factor;

        // Split into integer and fraction parts
        double int_part_d = std::floor(value + float_epsilon);
        long long frac_val = 0;
        if (total_frac > 0)
        {
            const double frac = value - int_part_d;
            // Epsilon added AFTER multiplication to avoid amplifying it by round_factor
            frac_val = static_cast<long long>(std::round(frac * round_factor + float_epsilon));
            // Handle rounding overflow (e.g., 0.999 → 1.000)
            if (frac_val >= static_cast<long long>(round_factor)) { int_part_d += 1.0; frac_val = 0; }
        }

        // Suppress negative sign when the rounded value is zero.
        // ECMA-376: negative values that round to zero display as "0" (not "-0").
        const bool is_zero_rounded = (int_part_d == 0.0 && frac_val == 0);
        const bool effective_negative = negative && !is_zero_rounded;

        // Format integer part as string (avoids long long overflow for large values)
        char int_buf[64];
        auto int_r = std::to_chars(int_buf, int_buf + sizeof(int_buf), int_part_d, std::chars_format::fixed, 0);
        const std::string int_str_raw(int_buf, int_r.ptr);

        // Format integer and fraction strings
        std::string int_str = format_integer_str(int_str_raw, dc);
        std::string frac_str = format_decimal_str(frac_val, total_frac, dc);

        // Walk tokens and assemble output
        std::string result;
        result.reserve(int_str.size() + frac_str.size() + 16);
        if (effective_negative && !neg_section_owns_sign) result += '-';

        bool past_decimal = false;
        bool int_output = false;
        int frac_pos = 0;

        for (auto& tok : sec.tokens)
        {
            switch (tok.type)
            {
            default:
                if (append_common_token(result, tok)) break;
                break;
            case token_type::decimal:
                // ECMA-376: the decimal point always appears in output when present
                result += '.';
                past_decimal = true;
                break;
            case token_type::thousands: break; // handled by format_integer_str
            case token_type::percent: result += '%'; break;
            case token_type::text_placeholder: result += '@'; break;
            case token_type::scientific:
                result += tok.exp_upper_case ? 'E' : 'e';
                result += tok.exp_plus_sign ? '+' : '-';
                result += tok.exp_pattern;
                break;
            case token_type::digit_zero: case token_type::digit_hash: case token_type::digit_qmark:
                if (!past_decimal)
                {
                    if (!int_output) { result += int_str; int_output = true; }
                }
                else
                {
                    if (frac_pos < static_cast<int>(frac_str.size())) { result += frac_str[frac_pos]; ++frac_pos; }
                    else if (tok.type == token_type::digit_qmark) result += ' '; // ? → space
                }
                break;
            }
        }

        // ECMA-376: when there are no integer digit placeholders but the integer
        // part is non-zero, the integer part is still displayed.
        if (!int_output && !int_str.empty() && int_part_d > 0)
        {
            // Insert before the decimal point (or at start if no decimal point).
            // E.g., "$".# with value 1.5 → "$1.5", not "1$.5".
            const size_t dot = result.find('.');
            const size_t ins = (dot != std::string::npos) ? dot : 0;
            result.insert(ins, int_str);
        }
        return result;
    }

private:
    // Count digit placeholders (0, #, ?) in a section, separated by the decimal point.
    // If end_pos is specified, only counts tokens before that position.
    // Records the L-to-R order of integer placeholders in dc.int_pattern for
    // correct zero-padding (e.g., "0#" → "05" vs "#0" → "5").
    //
    // Grouping vs. scaling commas: a comma is a grouping (thousands) separator
    // only if it has digit placeholders on BOTH sides (ECMA-376: "between digit
    // placeholders").  A comma to the left of the first digit or to the right
    // of the last digit is a scaling comma, handled separately by compute_scale.
    static digit_counts count_digit_placeholders(const section& sec, int end_pos = -1) noexcept
    {
        if (end_pos < 0) end_pos = static_cast<int>(sec.tokens.size());

        digit_counts dc;
        bool past_decimal = false;

        // Find the decimal point and the boundaries of integer digit placeholders.
        // Grouping commas must lie strictly between first and last integer digits.
        int decimal_pos = end_pos;
        int first_int_digit = -1;
        int last_int_digit = -1;
        for (int i = 0; i < end_pos; ++i)
            if (sec.tokens[i].type == token_type::decimal) { decimal_pos = i; break; }
        for (int i = 0; i < decimal_pos; ++i)
        {
            if (is_digit_token(sec.tokens[i].type))
            {
                if (first_int_digit < 0) first_int_digit = i;
                last_int_digit = i;
            }
        }

        for (int i = 0; i < end_pos; ++i)
        {
            const auto tt = sec.tokens[i].type;
            if (tt == token_type::decimal)       { past_decimal = true; continue; }
            if (tt == token_type::percent || tt == token_type::scientific) continue;
            if (tt == token_type::thousands)     {
                // Grouping comma: must have digit placeholders on BOTH sides
                // (i.e., strictly between first_int_digit and last_int_digit)
                if (!past_decimal && first_int_digit >= 0 && i > first_int_digit && i < last_int_digit)
                    ++dc.thousands;
                continue;
            }

            if (!past_decimal)
            {
                if (tt == token_type::digit_zero)       { ++dc.int_zeros;  dc.int_pattern += '0'; }
                else if (tt == token_type::digit_hash)  { ++dc.int_hashes; dc.int_pattern += '#'; }
                else if (tt == token_type::digit_qmark) { ++dc.int_qmarks; dc.int_pattern += '?'; }
            }
            else
            {
                if (tt == token_type::digit_zero)       { ++dc.frac_zeros;  dc.frac_pattern += '0'; }
                else if (tt == token_type::digit_hash)  { ++dc.frac_hashes; dc.frac_pattern += '#'; }
                else if (tt == token_type::digit_qmark) { ++dc.frac_qmarks; dc.frac_pattern += '?'; }
            }
        }
        return dc;
    }

    // Format the integer part of a regular number.
    // Handles: leading zero padding (0 placeholders), leading zero suppression
    // (# placeholders), space padding (? placeholders), thousands separators,
    // and zero suppression.
    //
    // Algorithm: pad the string to total_int, then walk the pattern L-to-R.
    // Before the first non-zero digit: '0'→'0', '#'→skip, '?'→' '.
    // At/after the first non-zero digit: output the actual digit.
    // This directly builds the correct output without fragile space-marking.
    static std::string format_integer_str(const std::string& int_str_raw, const digit_counts& dc)
    {
        // Handle zero value with no mandatory zero placeholders
        if (int_str_raw == "0")
        {
            if (dc.int_zeros == 0 && (dc.int_hashes > 0 || dc.int_qmarks > 0))
            {
                if (dc.int_qmarks > 0)
                    return std::string(dc.int_qmarks, ' ');
                return {};
            }
        }

        const int total = dc.total_int();
        std::string padded = int_str_raw;

        // Pad to the total number of integer placeholders
        if (static_cast<int>(padded.size()) < total)
            padded.insert(0, static_cast<size_t>(total - static_cast<int>(padded.size())), '0');

        // Find the first non-zero digit. If all zeros, treat all as "before".
        int first_nonzero = -1;
        for (size_t i = 0; i < padded.size(); ++i)
            if (padded[i] != '0') { first_nonzero = static_cast<int>(i); break; }
        if (first_nonzero < 0) first_nonzero = static_cast<int>(padded.size());

        // Build output directly from the pattern
        std::string result;
        result.reserve(padded.size() + dc.thousands);
        for (size_t i = 0; i < padded.size(); ++i)
        {
            const bool before = (static_cast<int>(i) < first_nonzero);
            const char pat = (i < dc.int_pattern.size()) ? dc.int_pattern[i] : '0';

            if (before)
            {
                if (pat == '#') continue;          // suppress
                if (pat == '?') result += ' ';      // space for alignment
                else result += '0';                  // '0' placeholder
            }
            else
            {
                result += padded[i];
            }
        }

        // Insert thousands separators every 3 digits from the right.
        // Only count actual digit characters (not spaces from ? padding).
        if (dc.thousands > 0)
        {
            int digit_count = 0;
            for (char ch : result)
                if (ch >= '0' && ch <= '9') ++digit_count;

            if (digit_count > 3)
            {
                std::string with_commas;
                with_commas.reserve(result.size() + (digit_count - 1) / 3);
                int digits_seen = 0;
                for (char ch : result)
                {
                    if (ch >= '0' && ch <= '9')
                    {
                        // Insert comma every 3 digits counting from right
                        if (digits_seen > 0 && (digit_count - digits_seen) % 3 == 0)
                            with_commas += ',';
                        ++digits_seen;
                    }
                    with_commas += ch;
                }
                result = std::move(with_commas);
            }
        }
        return result;
    }

    // Format the decimal part of a regular number.
    // Uses frac_pattern to correctly distinguish 0/#/? positions:
    //   0 → always show digit (even if zero)
    //   # → show digit, suppress trailing zeros entirely
    //   ? → show digit if non-zero, show space for zero (decimal alignment)
    //
    // Trailing suppression: positions beyond the rightmost "must-keep"
    // position are dropped.  A position must be kept if its placeholder is
    // '0' or '?', or if its digit is non-zero.
    static std::string format_decimal_str(long long frac_val, int total_frac, const digit_counts& dc)
    {
        if (total_frac <= 0) return {};

        // Convert to zero-padded string of length total_frac
        std::string s = std::to_string(frac_val);
        if (static_cast<int>(s.size()) < total_frac)
            s.insert(0, static_cast<size_t>(total_frac - static_cast<int>(s.size())), '0');

        // Find the rightmost position that must be kept.
        // '0' and '?' placeholders always keep their position (the latter
        // shows a space for zero).  Non-zero digits always keep their position.
        int last_keep = -1;
        for (int i = static_cast<int>(s.size()) - 1; i >= 0; --i)
        {
            const char pat = (i < static_cast<int>(dc.frac_pattern.size()))
                ? dc.frac_pattern[i] : '0';
            if (pat == '0' || pat == '?' || s[i] != '0') { last_keep = i; break; }
        }
        if (last_keep < 0) return {};

        // Build output with per-position placeholder rules
        std::string result;
        result.reserve(static_cast<size_t>(last_keep) + 1);
        for (int i = 0; i <= last_keep; ++i)
        {
            const char pat = (i < static_cast<int>(dc.frac_pattern.size()))
                ? dc.frac_pattern[i] : '0';
            if (s[i] != '0')
                result += s[i];           // non-zero digit → always show
            else if (pat == '?')
                result += ' ';            // ? placeholder → space for zero
            else
                result += '0';            // '0' or '#' placeholder → show zero
            // '#' with zero digit before last_keep is not a trailing zero
            // (it is followed by a mandatory '0' or '?' placeholder), so
            // it is correctly shown as '0' above.
        }
        return result;
    }
};

// =============================================================================
// number_format public API — thin wrappers delegating to _impl
// =============================================================================

number_format::number_format()
    : _impl(std::make_unique<impl>())
{}

number_format::number_format(const std::string& format_string)
    : _impl(std::make_unique<impl>())
{
    _impl->parse(format_string);
}

number_format::~number_format() = default;

number_format::number_format(number_format&&) noexcept = default;
number_format& number_format::operator=(number_format&&) noexcept = default;

std::string number_format::format(double number, bool date1904) const
{
    return _impl->format_double(number, date1904);
}

std::string number_format::format(const std::string& text) const
{
    return _impl->format_text(text);
}

} // namespace xlsxtext