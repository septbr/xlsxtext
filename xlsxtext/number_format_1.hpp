#pragma once

#include <cstring>
#include <string>
#include <vector>
#include <cmath>
#include <cstdint>
#include <charconv>

namespace xlsxtext
{

    class number_format
    {
    public:
        number_format() = default;
        number_format(const std::string &format_string) noexcept
        {
            parse(format_string);
        }

        std::string format(double number, bool date1904 = false) const;
        std::string format(const std::string &text) const;

    private:
        enum class token_type : uint8_t
        {
            literal,
            digit_zero,
            digit_hash,
            digit_qmark,
            decimal,
            thousands,
            percent,
            scientific,
            text_placeholder,
            skip,
            fill,
            year_2,
            year_4,
            month_n,
            month_nn,
            month_mmm,
            month_mmmm,
            day_d,
            day_dd,
            day_ddd,
            day_dddd,
            hour_h,
            hour_hh,
            minute_m,
            minute_mm,
            second_s,
            second_ss,
            am_pm,
            elapsed_hours,
            elapsed_minutes,
            elapsed_seconds,
            frac_second,
        };

        struct token
        {
            token_type type = token_type::literal;
            std::string literal;
            int repeat = 0;
            bool exp_plus_sign = true;
            int exp_digits = 0;
        };

        struct section
        {
            std::vector<token> tokens;
            bool has_condition = false;
            enum cond_op : uint8_t
            {
                cond_none,
                cond_gt,
                cond_ge,
                cond_lt,
                cond_le,
                cond_eq,
                cond_ne
            };
            cond_op condition_op = cond_none;
            double condition_value = 0;
            std::string color;
            int section_type = 0;
            int scale = 1;
        };

        struct date_parts
        {
            int year;
            int month;
            int day;
            int hour;
            int minute;
            int second;
            int dow;
            double fsec;
        };

        std::vector<section> _sections;
        bool _is_general = false;

        void parse(const std::string &fmt);

        static bool is_date_token(token_type t) noexcept;
        static bool is_time_token(token_type t) noexcept;

        const section *select_section(double number) const noexcept;
        std::string format_number_general(double number) const;
        std::string format_number_section(const section &sec, double raw_value,
                                          bool neg_section_owns_sign, bool date1904) const;

        static void best_fraction(double frac, int max_den, int &num, int &den) noexcept;

        static date_parts serial_to_date(double serial, bool date1904) noexcept;
        static int days_from_civil(int y, int m, int d) noexcept;
        static void civil_from_days(int z, int &y, int &m, int &d) noexcept;
    };

    // =========================================================================
    // Date helpers (Howard Hinnant civil calendar algorithms)
    // =========================================================================

    inline int number_format::days_from_civil(int y, int m, int d) noexcept
    {
        y -= m <= 2;
        const int era = (y >= 0 ? y : y - 399) / 400;
        const int yoe = static_cast<int>(y - era * 400);
        const int doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1;
        const int doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
        return era * 146097 + doe - 719468;
    }

    inline void number_format::civil_from_days(int z, int &y, int &m, int &d) noexcept
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

    inline number_format::date_parts number_format::serial_to_date(double serial, bool date1904) noexcept
    {
        const int whole = static_cast<int>(std::floor(serial));
        const double frac = serial - whole;

        if (!date1904 && whole == 60)
        {
            const int dow = (whole + 6) % 7;
            const double total_seconds = frac * 86400.0;
            const int hour = static_cast<int>(total_seconds) / 3600;
            const int minute = (static_cast<int>(total_seconds) % 3600) / 60;
            double fsec = total_seconds - hour * 3600.0 - minute * 60.0;
            const int second = static_cast<int>(std::floor(fsec));
            fsec -= second;
            return {1900, 2, 29, hour, minute, second, dow, fsec};
        }

        int real_days = whole;
        if (!date1904 && real_days > 60)
            --real_days;

        int epoch_days;
        if (date1904)
            epoch_days = days_from_civil(1904, 1, 1);
        else
            epoch_days = days_from_civil(1899, 12, 31);

        int y, m, d;
        civil_from_days(epoch_days + real_days, y, m, d);

        const int dow = (whole + 6) % 7;

        const double total_seconds = frac * 86400.0;
        const int hour = static_cast<int>(total_seconds) / 3600;
        const int minute = (static_cast<int>(total_seconds) % 3600) / 60;
        double fsec = total_seconds - hour * 3600.0 - minute * 60.0;
        const int second = static_cast<int>(std::floor(fsec));
        fsec -= second;

        return {y, m, d, hour, minute, second, dow, fsec};
    }

    inline void number_format::best_fraction(double frac, int max_den, int &best_num, int &best_den) noexcept
    {
        best_num = 0;
        best_den = 1;
        double best_err = 1.0;
        for (int den = 1; den <= max_den; ++den)
        {
            const int num = static_cast<int>(std::round(frac * den));
            if (num > den)
                continue;
            const double err = std::fabs(frac - static_cast<double>(num) / den);
            if (err < best_err - 1e-12)
            {
                best_err = err;
                best_num = num;
                best_den = den;
            }
        }
        if (best_num == best_den)
            best_num = 0;
        if (best_num == 0)
            best_den = 1;
    }

    // =========================================================================
    // Token type predicates
    // =========================================================================

    inline bool number_format::is_date_token(token_type t) noexcept
    {
        switch (t)
        {
        case token_type::year_2:
        case token_type::year_4:
        case token_type::month_n:
        case token_type::month_nn:
        case token_type::month_mmm:
        case token_type::month_mmmm:
        case token_type::day_d:
        case token_type::day_dd:
        case token_type::day_ddd:
        case token_type::day_dddd:
            return true;
        default:
            return false;
        }
    }

    inline bool number_format::is_time_token(token_type t) noexcept
    {
        switch (t)
        {
        case token_type::hour_h:
        case token_type::hour_hh:
        case token_type::minute_m:
        case token_type::minute_mm:
        case token_type::second_s:
        case token_type::second_ss:
        case token_type::am_pm:
        case token_type::elapsed_hours:
        case token_type::elapsed_minutes:
        case token_type::elapsed_seconds:
        case token_type::frac_second:
            return true;
        default:
            return false;
        }
    }

    // =========================================================================
    // Parse format string into sections per ECMA-376
    // =========================================================================

    inline void number_format::parse(const std::string &fmt)
    {
        _sections.clear();
        _is_general = false;

        if (fmt.empty())
        {
            _is_general = true;
            return;
        }

        std::string upper;
        upper.reserve(fmt.size());
        for (auto c : fmt)
            upper += static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
        if (upper == "GENERAL")
        {
            _is_general = true;
            return;
        }

        // Step 1: split into sections by ';' (respecting quotes and brackets)
        std::vector<std::string> raw_sections;
        std::string cur;
        bool in_quote = false;
        bool in_bracket = false;

        for (size_t i = 0; i < fmt.size(); ++i)
        {
            const char c = fmt[i];

            if (c == '"' && !in_bracket)
            {
                in_quote = !in_quote;
                cur += c;
                continue;
            }
            if (!in_quote && c == '[')
            {
                in_bracket = true;
                cur += c;
                continue;
            }
            if (!in_quote && c == ']')
            {
                in_bracket = false;
                cur += c;
                continue;
            }
            if (!in_quote && !in_bracket && c == ';')
            {
                raw_sections.push_back(cur);
                cur.clear();
                continue;
            }
            if (c == '\\' && i + 1 < fmt.size())
            {
                cur += c;
                cur += fmt[++i];
                continue;
            }
            cur += c;
        }
        raw_sections.push_back(cur);

        // Step 2: parse each raw section into tokens
        for (size_t si = 0; si < raw_sections.size(); ++si)
        {
            section sec;
            sec.section_type = static_cast<int>(si);

            const std::string &raw = raw_sections[si];
            size_t pos = 0;

            // Parse leading bracket: condition, color, or DBNum
            if (pos < raw.size() && raw[pos] == '[')
            {
                const size_t end = raw.find(']', pos);
                if (end != std::string::npos)
                {
                    const std::string bracket = raw.substr(pos + 1, end - pos - 1);

                    const bool is_elapsed = (bracket.size() == 1 &&
                                             (bracket[0] == 'h' || bracket[0] == 'H' ||
                                              bracket[0] == 'm' || bracket[0] == 'M' ||
                                              bracket[0] == 's' || bracket[0] == 'S'));
                    const bool is_locale = (!bracket.empty() && bracket[0] == '$');

                    if (!is_elapsed && !is_locale)
                    {
                        pos = end + 1;

                        if (!bracket.empty() &&
                            (bracket[0] == '>' || bracket[0] == '<' || bracket[0] == '='))
                        {
                            sec.has_condition = true;
                            const char *p = bracket.c_str();
                            if (p[0] == '>' && p[1] == '=')
                            {
                                sec.condition_op = section::cond_ge;
                                p += 2;
                            }
                            else if (p[0] == '<' && p[1] == '=')
                            {
                                sec.condition_op = section::cond_le;
                                p += 2;
                            }
                            else if (p[0] == '<' && p[1] == '>')
                            {
                                sec.condition_op = section::cond_ne;
                                p += 2;
                            }
                            else if (p[0] == '>')
                            {
                                sec.condition_op = section::cond_gt;
                                p += 1;
                            }
                            else if (p[0] == '<')
                            {
                                sec.condition_op = section::cond_lt;
                                p += 1;
                            }
                            else if (p[0] == '=')
                            {
                                sec.condition_op = section::cond_eq;
                                p += 1;
                            }
                            sec.condition_value = std::strtod(p, nullptr);
                        }
                        else
                        {
                            sec.color = bracket;
                        }
                    }
                }
            }

            // Tokenize the remaining content
            struct raw_token
            {
                token_type type;
                int repeat;
                std::string lit;
            };
            std::vector<raw_token> raw_tokens;

            while (pos < raw.size())
            {
                const char c = raw[pos];

                if (c == '"')
                {
                    const size_t end = raw.find('"', pos + 1);
                    if (end == std::string::npos)
                        break;
                    raw_tokens.push_back({token_type::literal, 0, raw.substr(pos + 1, end - pos - 1)});
                    pos = end + 1;
                    continue;
                }
                if (c == '\\' && pos + 1 < raw.size())
                {
                    raw_tokens.push_back({token_type::literal, 0, std::string(1, raw[pos + 1])});
                    pos += 2;
                    continue;
                }
                if (c == '_' && pos + 1 < raw.size())
                {
                    raw_tokens.push_back({token_type::skip, 0, std::string(1, raw[pos + 1])});
                    pos += 2;
                    continue;
                }
                if (c == '*' && pos + 1 < raw.size())
                {
                    raw_tokens.push_back({token_type::fill, 0, std::string(1, raw[pos + 1])});
                    pos += 2;
                    continue;
                }
                if (c == '@')
                {
                    raw_tokens.push_back({token_type::text_placeholder, 0, ""});
                    ++pos;
                    continue;
                }

                // Elapsed time: [h] [m] [s]
                if (pos + 2 < raw.size() && raw[pos] == '[' && raw[pos + 2] == ']')
                {
                    const char mid = static_cast<char>(std::tolower(static_cast<unsigned char>(raw[pos + 1])));
                    if (mid == 'h')
                    {
                        raw_tokens.push_back({token_type::elapsed_hours, 0, ""});
                        pos += 3;
                        continue;
                    }
                    if (mid == 'm')
                    {
                        raw_tokens.push_back({token_type::elapsed_minutes, 0, ""});
                        pos += 3;
                        continue;
                    }
                    if (mid == 's')
                    {
                        raw_tokens.push_back({token_type::elapsed_seconds, 0, ""});
                        pos += 3;
                        continue;
                    }
                }

                // Scientific notation: E+ E- e+ e-
                if ((c == 'E' || c == 'e') && pos + 1 < raw.size() &&
                    (raw[pos + 1] == '+' || raw[pos + 1] == '-'))
                {
                    const bool plus = (raw[pos + 1] == '+');
                    pos += 2;
                    int edigits = 0;
                    while (pos < raw.size() && raw[pos] == '0')
                    {
                        ++edigits;
                        ++pos;
                    }
                    if (edigits == 0)
                        edigits = 1;
                    raw_tokens.push_back({token_type::scientific, edigits, plus ? "+" : "-"});
                    continue;
                }

                // Year: y/yy/yyy/yyyy
                if (c == 'y' || c == 'Y')
                {
                    int n = 1;
                    while (pos + n < raw.size() && (raw[pos + n] == 'y' || raw[pos + n] == 'Y'))
                        ++n;
                    raw_tokens.push_back({n >= 4 ? token_type::year_4 : token_type::year_2, n, ""});
                    pos += n;
                    continue;
                }

                // Month: m/mm/mmm/mmmm/mmmmm
                if (c == 'm' || c == 'M')
                {
                    int n = 1;
                    while (pos + n < raw.size() && (raw[pos + n] == 'm' || raw[pos + n] == 'M'))
                        ++n;
                    token_type t;
                    if (n >= 5)
                        t = token_type::month_n;
                    else if (n == 4)
                        t = token_type::month_mmmm;
                    else if (n == 3)
                        t = token_type::month_mmm;
                    else
                        t = (n == 2) ? token_type::month_nn : token_type::month_n;
                    raw_tokens.push_back({t, n, ""});
                    pos += n;
                    continue;
                }

                // Day: d/dd/ddd/dddd
                if (c == 'd' || c == 'D')
                {
                    int n = 1;
                    while (pos + n < raw.size() && (raw[pos + n] == 'd' || raw[pos + n] == 'D'))
                        ++n;
                    token_type t;
                    if (n >= 4)
                        t = token_type::day_dddd;
                    else if (n == 3)
                        t = token_type::day_ddd;
                    else if (n == 2)
                        t = token_type::day_dd;
                    else
                        t = token_type::day_d;
                    raw_tokens.push_back({t, n, ""});
                    pos += n;
                    continue;
                }

                // Hour: h/hh
                if (c == 'h' || c == 'H')
                {
                    int n = 1;
                    while (pos + n < raw.size() && (raw[pos + n] == 'h' || raw[pos + n] == 'H'))
                        ++n;
                    raw_tokens.push_back({n >= 2 ? token_type::hour_hh : token_type::hour_h, n, ""});
                    pos += n;
                    continue;
                }

                // Second: s/ss
                if (c == 's' || c == 'S')
                {
                    int n = 1;
                    while (pos + n < raw.size() && (raw[pos + n] == 's' || raw[pos + n] == 'S'))
                        ++n;
                    raw_tokens.push_back({n >= 2 ? token_type::second_ss : token_type::second_s, n, ""});
                    pos += n;
                    continue;
                }

                // AM/PM
                if (pos + 4 < raw.size())
                {
                    const std::string ampm = raw.substr(pos, 5);
                    if (ampm == "AM/PM" || ampm == "am/pm")
                    {
                        raw_tokens.push_back({token_type::am_pm, 0, ampm});
                        pos += 5;
                        continue;
                    }
                }
                if (pos + 2 < raw.size())
                {
                    const std::string ap = raw.substr(pos, 3);
                    if (ap == "A/P" || ap == "a/p")
                    {
                        raw_tokens.push_back({token_type::am_pm, 0, ap});
                        pos += 3;
                        continue;
                    }
                }

                // Decimal point
                if (c == '.')
                {
                    raw_tokens.push_back({token_type::decimal, 0, ""});
                    ++pos;
                    continue;
                }

                // Percent
                if (c == '%')
                {
                    raw_tokens.push_back({token_type::percent, 0, ""});
                    ++pos;
                    continue;
                }

                // Comma (thousands separator / scaling)
                if (c == ',')
                {
                    raw_tokens.push_back({token_type::thousands, 0, ""});
                    ++pos;
                    continue;
                }

                // Colon (time separator)
                if (c == ':')
                {
                    raw_tokens.push_back({token_type::literal, 0, ":"});
                    ++pos;
                    continue;
                }

                // Slash (literal – fraction detected later)
                if (c == '/')
                {
                    raw_tokens.push_back({token_type::literal, 0, "/"});
                    ++pos;
                    continue;
                }

                // 0 digit placeholder
                if (c == '0')
                {
                    int n = 1;
                    while (pos + n < raw.size() && raw[pos + n] == '0')
                        ++n;
                    raw_tokens.push_back({token_type::digit_zero, n, ""});
                    pos += n;
                    continue;
                }

                // # digit placeholder
                if (c == '#')
                {
                    int n = 1;
                    while (pos + n < raw.size() && raw[pos + n] == '#')
                        ++n;
                    raw_tokens.push_back({token_type::digit_hash, n, ""});
                    pos += n;
                    continue;
                }

                // ? digit placeholder
                if (c == '?')
                {
                    int n = 1;
                    while (pos + n < raw.size() && raw[pos + n] == '?')
                        ++n;
                    raw_tokens.push_back({token_type::digit_qmark, n, ""});
                    pos += n;
                    continue;
                }

                // Any other character – literal
                raw_tokens.push_back({token_type::literal, 0, std::string(1, c)});
                ++pos;
            }

            // Resolve ambiguous m/mm tokens (month vs minute)
            // Per ECMA-376: m/mm after h/hh or before s/ss → minute
            //                m/mm with h present and no d/y → minute
            bool has_h = false, has_d = false, has_y = false;
            for (auto &rt : raw_tokens)
            {
                if (rt.type == token_type::hour_h || rt.type == token_type::hour_hh)
                    has_h = true;
                if (rt.type == token_type::day_d || rt.type == token_type::day_dd ||
                    rt.type == token_type::day_ddd || rt.type == token_type::day_dddd)
                    has_d = true;
                if (rt.type == token_type::year_2 || rt.type == token_type::year_4)
                    has_y = true;
            }

            for (size_t i = 0; i < raw_tokens.size(); ++i)
            {
                auto &rt = raw_tokens[i];
                if (rt.type != token_type::month_n && rt.type != token_type::month_nn)
                    continue;

                const bool after_h = (i > 0 &&
                                      (raw_tokens[i - 1].type == token_type::hour_h ||
                                       raw_tokens[i - 1].type == token_type::hour_hh));
                const bool before_s = (i + 1 < raw_tokens.size() &&
                                       (raw_tokens[i + 1].type == token_type::second_s ||
                                        raw_tokens[i + 1].type == token_type::second_ss ||
                                        raw_tokens[i + 1].type == token_type::frac_second));
                const bool colon_before = (i > 0 &&
                                           raw_tokens[i - 1].type == token_type::literal &&
                                           raw_tokens[i - 1].lit == ":");
                const bool colon_after = (i + 1 < raw_tokens.size() &&
                                          raw_tokens[i + 1].type == token_type::literal &&
                                          raw_tokens[i + 1].lit == ":");

                if (after_h || before_s || colon_before || colon_after ||
                    (has_h && !has_d && !has_y))
                {
                    rt.type = (rt.type == token_type::month_nn)
                                  ? token_type::minute_mm
                                  : token_type::minute_m;
                }
            }

            // Detect fraction seconds: .0 .00 .000 after s/ss
            for (size_t i = 0; i < raw_tokens.size(); ++i)
            {
                if (raw_tokens[i].type == token_type::decimal &&
                    i > 0 && (raw_tokens[i - 1].type == token_type::second_s || raw_tokens[i - 1].type == token_type::second_ss))
                {
                    int nz = 0;
                    size_t j = i + 1;
                    while (j < raw_tokens.size() && raw_tokens[j].type == token_type::digit_zero)
                    {
                        nz += raw_tokens[j].repeat;
                        ++j;
                    }
                    if (nz > 0)
                    {
                        raw_tokens[i].type = token_type::frac_second;
                        raw_tokens[i].repeat = nz;
                        for (size_t k = i + 1; k < j; ++k)
                            raw_tokens[k].type = token_type::literal;
                    }
                }
            }

            // Build final token list, flattening grouped digit placeholders
            for (auto &rt : raw_tokens)
            {
                if (rt.type == token_type::literal && rt.lit.empty())
                    continue;

                const bool is_digit = (rt.type == token_type::digit_zero ||
                                       rt.type == token_type::digit_hash ||
                                       rt.type == token_type::digit_qmark);
                const int count = is_digit ? rt.repeat : 1;

                for (int r = 0; r < count; ++r)
                {
                    token t;
                    t.type = rt.type;
                    t.repeat = (rt.type == token_type::frac_second) ? rt.repeat : 1;

                    if (rt.type == token_type::literal || rt.type == token_type::skip ||
                        rt.type == token_type::fill || rt.type == token_type::am_pm)
                        t.literal = rt.lit;

                    if (rt.type == token_type::scientific)
                    {
                        t.exp_plus_sign = (rt.lit == "+");
                        t.exp_digits = rt.repeat;
                    }

                    sec.tokens.push_back(t);
                }
            }

            // Compute scale from trailing commas
            sec.scale = 1;
            int last_digit_pos = -1;
            for (size_t i = 0; i < sec.tokens.size(); ++i)
            {
                const auto tt = sec.tokens[i].type;
                if (tt == token_type::digit_zero || tt == token_type::digit_hash ||
                    tt == token_type::digit_qmark)
                    last_digit_pos = static_cast<int>(i);
            }
            for (size_t i = static_cast<size_t>(last_digit_pos + 1); i < sec.tokens.size(); ++i)
            {
                if (sec.tokens[i].type == token_type::thousands)
                    sec.scale *= 1000;
                else
                    break;
            }

            _sections.push_back(std::move(sec));
        }

        if (_sections.empty())
        {
            section s;
            s.tokens.push_back({token_type::literal, "", 0});
            _sections.push_back(std::move(s));
        }
    }

    // =========================================================================
    // Select the right section for a number per ECMA-376 section rules
    // =========================================================================

    inline const number_format::section *number_format::select_section(double number) const noexcept
    {
        if (_sections.empty())
            return nullptr;

        // Check if any section has conditions
        bool any_condition = false;
        for (auto &sec : _sections)
        {
            if (sec.has_condition)
            {
                any_condition = true;
                break;
            }
        }

        // Check conditional sections first
        for (auto &sec : _sections)
        {
            if (!sec.has_condition)
                continue;

            bool match = false;
            switch (sec.condition_op)
            {
            case section::cond_gt:
                match = number > sec.condition_value;
                break;
            case section::cond_ge:
                match = number >= sec.condition_value;
                break;
            case section::cond_lt:
                match = number < sec.condition_value;
                break;
            case section::cond_le:
                match = number <= sec.condition_value;
                break;
            case section::cond_eq:
                match = number == sec.condition_value;
                break;
            case section::cond_ne:
                match = number != sec.condition_value;
                break;
            default:
                break;
            }
            if (match)
                return &sec;
        }

        // If conditions exist but none matched, use first non-conditional section
        if (any_condition)
        {
            for (auto &sec : _sections)
                if (!sec.has_condition)
                    return &sec;
        }

        // Per ECMA-376:
        // 1 section:  applies to all numbers
        // 2 sections: first = positive+zero, second = negative
        // 3 sections: first = positive, second = negative, third = zero
        // 4 sections: first = positive, second = negative, third = zero, fourth = text

        const size_t n = _sections.size();

        if (n == 1)
        {
            return &_sections[0];
        }
        else if (n == 2)
        {
            return (number < 0) ? &_sections[1] : &_sections[0];
        }
        else // n >= 3
        {
            if (number > 0)
                return &_sections[0];
            else if (number < 0)
                return &_sections[1];
            else // zero
                return &_sections[2];
        }
    }

    inline std::string number_format::format_number_general(double number) const
    {
        char buf[64]{};
        std::to_chars(buf, buf + sizeof(buf), number);
        return std::string(buf);
    }

    // =========================================================================
    // Format number using a section
    // =========================================================================

    inline std::string number_format::format_number_section(const section &sec, double raw_value,
                                                            bool neg_section_owns_sign, bool date1904) const
    {
        // Check if text-only section (no digit placeholders, no date/time tokens)
        bool has_digits = false;
        bool has_date_time = false;
        for (auto &tok : sec.tokens)
        {
            if (tok.type == token_type::digit_zero || tok.type == token_type::digit_hash ||
                tok.type == token_type::digit_qmark)
                has_digits = true;
            if (is_date_token(tok.type) || is_time_token(tok.type))
                has_date_time = true;
        }

        if (!has_digits && !has_date_time)
        {
            std::string result;
            for (auto &tok : sec.tokens)
            {
                if (tok.type == token_type::literal)
                    result += tok.literal;
                else if (tok.type == token_type::text_placeholder)
                    result += format_number_general(raw_value);
                else if (tok.type == token_type::skip)
                    result += ' ';
            }
            return result;
        }

        // Check section type
        bool has_date = false;
        bool has_time = false;
        for (auto &tok : sec.tokens)
        {
            if (is_date_token(tok.type))
                has_date = true;
            if (is_time_token(tok.type))
                has_time = true;
        }

        // Detect fraction format
        bool is_fraction = false;
        int frac_nd = 0;
        int frac_dd = 0;
        for (size_t i = 1; i + 1 < sec.tokens.size(); ++i)
        {
            if (sec.tokens[i].type == token_type::literal && sec.tokens[i].literal == "/")
            {
                int nd = 0, dd = 0;
                for (int j = static_cast<int>(i) - 1; j >= 0; --j)
                {
                    const auto tt = sec.tokens[j].type;
                    if (tt == token_type::digit_zero || tt == token_type::digit_hash ||
                        tt == token_type::digit_qmark)
                        nd += sec.tokens[j].repeat;
                    else if (tt == token_type::literal && sec.tokens[j].literal == " ")
                        break;
                    else
                        break;
                }
                for (size_t j = i + 1; j < sec.tokens.size(); ++j)
                {
                    const auto tt = sec.tokens[j].type;
                    if (tt == token_type::digit_zero || tt == token_type::digit_hash ||
                        tt == token_type::digit_qmark)
                        dd += sec.tokens[j].repeat;
                    else
                        break;
                }
                if (nd > 0 && dd > 0)
                {
                    is_fraction = true;
                    frac_nd = nd;
                    frac_dd = dd;
                }
                break;
            }
        }

        // Date/Time formatting
        if (has_date || has_time)
        {
            bool has_ampm = false;
            for (auto &tok : sec.tokens)
            {
                if (tok.type == token_type::am_pm)
                {
                    has_ampm = true;
                    break;
                }
            }

            const date_parts dp = serial_to_date(raw_value, date1904);
            std::string result;

            for (auto &tok : sec.tokens)
            {
                switch (tok.type)
                {
                case token_type::year_2:
                {
                    const int yy = dp.year % 100;
                    if (yy < 10)
                        result += '0';
                    result += std::to_string(yy);
                    break;
                }
                case token_type::year_4:
                {
                    const int y = dp.year;
                    if (y < 1000)
                        result += '0';
                    if (y < 100)
                        result += '0';
                    if (y < 10)
                        result += '0';
                    result += std::to_string(y);
                    break;
                }
                case token_type::month_n:
                    result += std::to_string(dp.month);
                    break;
                case token_type::month_nn:
                    if (dp.month < 10)
                        result += '0';
                    result += std::to_string(dp.month);
                    break;
                case token_type::month_mmm:
                {
                    static const char *m3[] = {
                        "", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
                    result += m3[dp.month];
                    break;
                }
                case token_type::month_mmmm:
                {
                    static const char *m4[] = {
                        "", "January", "February", "March", "April", "May", "June",
                        "July", "August", "September", "October", "November", "December"};
                    result += m4[dp.month];
                    break;
                }
                case token_type::day_d:
                    result += std::to_string(dp.day);
                    break;
                case token_type::day_dd:
                    if (dp.day < 10)
                        result += '0';
                    result += std::to_string(dp.day);
                    break;
                case token_type::day_ddd:
                {
                    static const char *d3[] = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"};
                    result += d3[dp.dow];
                    break;
                }
                case token_type::day_dddd:
                {
                    static const char *d4[] = {
                        "Sunday", "Monday", "Tuesday", "Wednesday",
                        "Thursday", "Friday", "Saturday"};
                    result += d4[dp.dow];
                    break;
                }
                case token_type::hour_h:
                {
                    if (has_ampm)
                    {
                        int h = dp.hour % 12;
                        if (h == 0)
                            h = 12;
                        result += std::to_string(h);
                    }
                    else
                    {
                        result += std::to_string(dp.hour);
                    }
                    break;
                }
                case token_type::hour_hh:
                {
                    if (has_ampm)
                    {
                        int h = dp.hour % 12;
                        if (h == 0)
                            h = 12;
                        if (h < 10)
                            result += '0';
                        result += std::to_string(h);
                    }
                    else
                    {
                        if (dp.hour < 10)
                            result += '0';
                        result += std::to_string(dp.hour);
                    }
                    break;
                }
                case token_type::minute_m:
                    result += std::to_string(dp.minute);
                    break;
                case token_type::minute_mm:
                    if (dp.minute < 10)
                        result += '0';
                    result += std::to_string(dp.minute);
                    break;
                case token_type::second_s:
                    result += std::to_string(dp.second);
                    break;
                case token_type::second_ss:
                    if (dp.second < 10)
                        result += '0';
                    result += std::to_string(dp.second);
                    break;
                case token_type::am_pm:
                {
                    const bool pm = dp.hour >= 12;
                    const auto &lit = tok.literal;
                    if (lit.size() == 5)
                        result += pm ? (lit[0] == 'A' ? "PM" : "pm") : (lit[0] == 'A' ? "AM" : "am");
                    else
                        result += pm ? (lit[0] == 'A' ? "P" : "p") : (lit[0] == 'A' ? "A" : "a");
                    break;
                }
                case token_type::frac_second:
                {
                    const double fs = dp.fsec;
                    const double pow10 = std::pow(10.0, tok.repeat);
                    int frac_val = static_cast<int>(std::round(fs * pow10));
                    if (frac_val >= static_cast<int>(pow10))
                        frac_val = static_cast<int>(pow10) - 1;
                    std::string fs_str = std::to_string(frac_val);
                    while (static_cast<int>(fs_str.size()) < tok.repeat)
                        fs_str = "0" + fs_str;
                    result += "." + fs_str;
                    break;
                }
                case token_type::elapsed_hours:
                    result += std::to_string(static_cast<int>(std::floor(raw_value * 24)));
                    break;
                case token_type::elapsed_minutes:
                    result += std::to_string(static_cast<int>(std::floor(raw_value * 24 * 60)));
                    break;
                case token_type::elapsed_seconds:
                    result += std::to_string(static_cast<int>(std::floor(raw_value * 24 * 3600)));
                    break;
                case token_type::literal:
                    result += tok.literal;
                    break;
                case token_type::skip:
                    result += ' ';
                    break;
                case token_type::fill:
                    result += tok.literal;
                    break;
                default:
                    break;
                }
            }
            return result;
        }

        // Fraction formatting
        if (is_fraction)
        {
            double scaled = raw_value / sec.scale;
            const bool negative = scaled < 0;
            if (negative)
                scaled = -scaled;

            bool has_percent = false;
            for (auto &tok : sec.tokens)
            {
                if (tok.type == token_type::percent)
                {
                    has_percent = true;
                    break;
                }
            }
            if (has_percent)
                scaled *= 100.0;

            int int_part = static_cast<int>(std::floor(scaled));
            const double frac_part = scaled - int_part;

            int max_den = 1;
            for (int i = 0; i < frac_dd; ++i)
                max_den *= 10;
            --max_den;

            int best_num, best_den;
            best_fraction(frac_part, max_den, best_num, best_den);

            if (best_num == best_den)
            {
                ++int_part;
                best_num = 0;
                best_den = 1;
            }

            std::string result;
            if (negative && !neg_section_owns_sign)
                result += '-';

            // Detect if format has explicit integer space
            bool has_int = false;
            for (auto &tok : sec.tokens)
            {
                if (tok.type == token_type::digit_zero || tok.type == token_type::digit_hash ||
                    tok.type == token_type::digit_qmark)
                {
                    has_int = true;
                    break;
                }
                if (tok.type == token_type::literal && tok.literal == "/")
                    break;
                if (tok.type == token_type::literal && tok.literal == " ")
                {
                    has_int = true;
                    break;
                }
            }

            if (has_int || int_part > 0)
                result += std::to_string(int_part);

            if (best_num > 0 || best_den > 1)
            {
                if (has_int || int_part > 0)
                    result += ' ';
                result += std::to_string(best_num);
                result += '/';
                result += std::to_string(best_den);
            }
            return result;
        }

        // Regular number formatting
        double value = raw_value / sec.scale;
        const bool negative = value < 0;
        if (negative)
            value = -value;

        bool has_percent = false;
        for (auto &tok : sec.tokens)
        {
            if (tok.type == token_type::percent)
            {
                has_percent = true;
                break;
            }
        }
        if (has_percent)
            value *= 100.0;

        // Scientific notation
        bool has_scientific = false;
        token sci_tok;
        int sci_pos = -1;
        for (size_t i = 0; i < sec.tokens.size(); ++i)
        {
            if (sec.tokens[i].type == token_type::scientific)
            {
                has_scientific = true;
                sci_tok = sec.tokens[i];
                sci_pos = static_cast<int>(i);
                break;
            }
        }

        if (has_scientific)
        {
            int exponent = 0;
            double mantissa = value;

            if (mantissa != 0.0)
            {
                exponent = static_cast<int>(std::floor(std::log10(mantissa)));
                mantissa = mantissa / std::pow(10.0, exponent);
            }

            int int_zeros = 0, int_hashes = 0;
            int frac_zeros = 0, frac_hashes = 0;
            bool past_decimal = false;

            for (int i = 0; i < sci_pos; ++i)
            {
                const auto tt = sec.tokens[i].type;
                if (tt == token_type::decimal)
                {
                    past_decimal = true;
                    continue;
                }
                if (!past_decimal)
                {
                    if (tt == token_type::digit_zero)
                        int_zeros += sec.tokens[i].repeat;
                    if (tt == token_type::digit_hash)
                        int_hashes += sec.tokens[i].repeat;
                }
                else
                {
                    if (tt == token_type::digit_zero)
                        frac_zeros += sec.tokens[i].repeat;
                    if (tt == token_type::digit_hash)
                        frac_hashes += sec.tokens[i].repeat;
                }
            }

            const int total_frac_digits = frac_zeros + frac_hashes;
            const double round_factor = std::pow(10.0, total_frac_digits);
            mantissa = std::round(mantissa * round_factor) / round_factor;

            if (mantissa >= 10.0 - 1e-12)
            {
                mantissa = 1.0;
                exponent += 1;
            }

            const int m_int = static_cast<int>(std::floor(mantissa));
            std::string m_int_str = std::to_string(m_int);
            while (static_cast<int>(m_int_str.size()) < int_zeros)
                m_int_str = "0" + m_int_str;

            std::string m_frac_str;
            if (total_frac_digits > 0)
            {
                const double m_frac = mantissa - m_int;
                const int frac_val = static_cast<int>(std::round(m_frac * round_factor));
                m_frac_str = std::to_string(frac_val);
                while (static_cast<int>(m_frac_str.size()) < total_frac_digits)
                    m_frac_str = "0" + m_frac_str;
                // Trim trailing hashes
                const int min_frac = frac_zeros;
                while (static_cast<int>(m_frac_str.size()) > min_frac &&
                       !m_frac_str.empty() && m_frac_str.back() == '0')
                    m_frac_str.pop_back();
            }

            std::string mantissa_result = m_int_str;
            if (total_frac_digits > 0 && !m_frac_str.empty())
                mantissa_result += "." + m_frac_str;

            std::string exp_str = std::to_string(std::abs(exponent));
            if (sci_tok.exp_digits > 0)
            {
                while (static_cast<int>(exp_str.size()) < sci_tok.exp_digits)
                    exp_str = "0" + exp_str;
            }

            std::string sign;
            if (exponent >= 0)
            {
                if (sci_tok.exp_plus_sign)
                    sign = "+";
            }
            else
            {
                sign = "-";
            }

            std::string result;
            if (negative && !neg_section_owns_sign)
                result += '-';
            result += mantissa_result;
            result += 'E';
            result += sign;
            result += exp_str;
            return result;
        }

        // Parse digit structure for regular number formatting
        int int_zeros = 0, int_hashes = 0, int_qmarks = 0;
        int frac_zeros = 0, frac_hashes = 0, frac_qmarks = 0;
        bool past_decimal = false;
        int thousands_count = 0;

        for (auto &tok : sec.tokens)
        {
            if (tok.type == token_type::decimal)
            {
                past_decimal = true;
                continue;
            }
            if (tok.type == token_type::percent || tok.type == token_type::scientific)
                continue;
            if (tok.type == token_type::thousands)
            {
                if (!past_decimal)
                    ++thousands_count;
                continue;
            }
            if (!past_decimal)
            {
                if (tok.type == token_type::digit_zero)
                    int_zeros += tok.repeat;
                else if (tok.type == token_type::digit_hash)
                    int_hashes += tok.repeat;
                else if (tok.type == token_type::digit_qmark)
                    int_qmarks += tok.repeat;
            }
            else
            {
                if (tok.type == token_type::digit_zero)
                    frac_zeros += tok.repeat;
                else if (tok.type == token_type::digit_hash)
                    frac_hashes += tok.repeat;
                else if (tok.type == token_type::digit_qmark)
                    frac_qmarks += tok.repeat;
            }
        }

        const int total_int_digits = int_zeros + int_hashes + int_qmarks;
        const int total_frac_digits = frac_zeros + frac_hashes + frac_qmarks;

        // Round to required precision
        const double round_factor = std::pow(10.0, total_frac_digits);
        value = std::round(value * round_factor) / round_factor;

        // Split into integer and fractional parts
        long long int_part = static_cast<long long>(std::floor(value + 1e-12));
        long long frac_val = 0;

        if (total_frac_digits > 0)
        {
            const double frac = value - int_part;
            frac_val = static_cast<long long>(std::round(frac * round_factor));
            if (frac_val >= static_cast<long long>(round_factor))
            {
                ++int_part;
                frac_val = 0;
            }
        }

        // Format integer part
        std::string int_str = std::to_string(std::llabs(int_part));

        // Pad with forced zeros
        while (static_cast<int>(int_str.size()) < int_zeros)
            int_str = "0" + int_str;

        const bool use_thousands = (thousands_count > 0);

        // Suppress leading zero for #-only formats per ECMA-376
        // # does not display extra zeros; if only # placeholders and no forced zeros,
        // a leading zero is non-significant and should be suppressed
        bool suppress_leading_zero = false;
        if (int_zeros == 0 && int_hashes > 0 && int_part == 0 && total_frac_digits > 0)
            suppress_leading_zero = true;

        if (suppress_leading_zero)
            int_str.clear();
        else if (int_zeros == 0 && int_str == "0" && total_frac_digits == 0 && !use_thousands)
            int_str.clear();

        // Apply thousands separator
        if (use_thousands && int_str.size() > 3)
        {
            std::string with_commas;
            const int len = static_cast<int>(int_str.size());
            for (int i = 0; i < len; ++i)
            {
                if (i > 0 && (len - i) % 3 == 0)
                    with_commas += ',';
                with_commas += int_str[i];
            }
            int_str = with_commas;
        }

        // Format fractional part
        std::string frac_str;
        if (total_frac_digits > 0)
        {
            frac_str = std::to_string(frac_val);
            while (static_cast<int>(frac_str.size()) < total_frac_digits)
                frac_str = "0" + frac_str;

            // Trim trailing # and ? placeholders (insignificant zeros)
            const int min_digits = (frac_qmarks > 0) ? 0 : frac_zeros;
            while (static_cast<int>(frac_str.size()) > min_digits &&
                   !frac_str.empty() && frac_str.back() == '0')
            {
                const int pos_from_right = static_cast<int>(frac_str.size());
                const int idx = total_frac_digits - pos_from_right;
                if (idx >= frac_zeros)
                    frac_str.pop_back();
                else
                    break;
            }

            // Pad ? placeholders with spaces for alignment
            if (frac_qmarks > 0)
            {
                while (static_cast<int>(frac_str.size()) < total_frac_digits)
                    frac_str += ' ';
            }
        }

        // Build output by walking through tokens
        std::string result;
        if (negative && !neg_section_owns_sign)
            result += '-';

        int frac_pos = 0;
        past_decimal = false;
        bool int_output = false;

        for (auto &tok : sec.tokens)
        {
            switch (tok.type)
            {
            case token_type::literal:
                result += tok.literal;
                break;
            case token_type::skip:
                result += ' ';
                break;
            case token_type::fill:
                result += tok.literal;
                break;
            case token_type::decimal:
                if (total_frac_digits > 0 && (!frac_str.empty() || frac_zeros > 0))
                    result += '.';
                past_decimal = true;
                break;
            case token_type::thousands:
                break;
            case token_type::percent:
                result += '%';
                break;
            case token_type::digit_zero:
            case token_type::digit_hash:
            case token_type::digit_qmark:
                if (!past_decimal)
                {
                    if (!int_output)
                    {
                        result += int_str;
                        int_output = true;
                    }
                }
                else
                {
                    if (frac_pos < static_cast<int>(frac_str.size()))
                    {
                        result += frac_str[frac_pos];
                        ++frac_pos;
                    }
                    else if (tok.type == token_type::digit_qmark)
                    {
                        result += ' ';
                    }
                }
                break;
            default:
                break;
            }
        }

        // Ensure integer part is present
        if (!int_output && !int_str.empty())
        {
            const size_t ins = (negative && !result.empty() && result[0] == '-') ? 1 : 0;
            result.insert(ins, int_str);
        }

        return result;
    }

    // =========================================================================
    // Public format methods
    // =========================================================================

    inline std::string number_format::format(double number, bool date1904) const
    {
        if (_is_general || _sections.empty())
            return format_number_general(number);

        const section *sec = select_section(number);
        if (!sec)
            return format_number_general(number);

        // Text section (section_type == 3): numbers rendered with @ placeholder
        if (sec->section_type == 3)
        {
            std::string result;
            for (auto &tok : sec->tokens)
            {
                if (tok.type == token_type::text_placeholder)
                    result += format_number_general(number);
                else if (tok.type == token_type::literal)
                    result += tok.literal;
                else if (tok.type == token_type::skip)
                    result += ' ';
            }
            return result;
        }

        // Determine if negative section already handles its own sign
        bool neg_section_owns_sign = false;
        if (number < 0 && _sections.size() >= 2)
        {
            if (sec == &_sections[1])
            {
                bool has_sign = false;
                for (auto &tok : sec->tokens)
                {
                    if (tok.type == token_type::literal)
                    {
                        if (tok.literal == "-" || tok.literal == "(")
                        {
                            has_sign = true;
                            break;
                        }
                    }
                    else
                        break;
                }
                neg_section_owns_sign = has_sign;
            }
        }

        return format_number_section(*sec, number, neg_section_owns_sign, date1904);
    }

    inline std::string number_format::format(const std::string &text) const
    {
        if (_is_general || _sections.empty())
            return text;

        // Find text section (section_type == 3, or last if >= 4)
        const section *text_sec = nullptr;
        for (auto &sec : _sections)
        {
            if (sec.section_type == 3)
            {
                text_sec = &sec;
                break;
            }
        }

        if (!text_sec && _sections.size() >= 4)
            text_sec = &_sections[3];
        else if (!text_sec)
            text_sec = &_sections[0];

        std::string result;
        for (auto &tok : text_sec->tokens)
        {
            if (tok.type == token_type::text_placeholder)
                result += text;
            else if (tok.type == token_type::literal)
                result += tok.literal;
            else if (tok.type == token_type::skip)
                result += ' ';
        }
        return result;
    }

} // namespace xlsxtext