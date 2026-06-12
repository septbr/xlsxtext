#pragma once

#include <cstring>
#include <string>
#include <vector>
#include <array>
#include <cmath>
#include <cstdint>
#include <stdexcept>

namespace xlsxtext
{
    class number_format
    {
    private:
        // Constants for number formatting
        static constexpr double general_small_threshold = 1e-4;
        static constexpr double general_large_threshold = 1e15;
        static constexpr double float_epsilon = 1e-10;
        static constexpr std::size_t default_cell_width = 11;

        struct datetime
        {
            int year;
            int month;
            int day;
            int hour;
            int minute;
            int second;
            int microsecond;

            datetime(int year_, int month_, int day_, int hour_ = 0, int minute_ = 0, int second_ = 0, int microsecond_ = 0)
                : year(year_), month(month_), day(day_), hour(hour_), minute(minute_), second(second_), microsecond(microsecond_)
            {
            }

            int weekday() const
            {
                int d = day, m = month, y = year;
                if (m == 1 || m == 2)
                {
                    m += 12;
                    y--;
                }
                int week = (d + 2 * m + 3 * (m + 1) / 5 + y + y / 4 - y / 100 + y / 400 + 1) % 7;
                return week;
            }

            static datetime from_number(double number, bool is_date1904 = false)
            {
                datetime result(0, 1, 0);

                int serial = static_cast<int>(number);

                if (is_date1904)
                    serial += 1462;

                // Excel 1900 date system: serial 1 = 1900-01-01
                // Excel bug: 1900-02-29 (serial 60) is fictional
                // For serial > 60, Excel dates are 1 day ahead of reality

                if (serial == 60)
                {
                    // Excel's fictional 1900-02-29
                    result.year = 1900;
                    result.month = 2;
                    result.day = 29;
                }
                else
                {
                    // Calculate Julian Day Number
                    // JDN for 1899-12-30 (Excel epoch for serial 0) = 2415020
                    // For serial > 60, subtract 1 to compensate for the bug
                    int jdn = serial + 2415020;
                    if (serial > 60)
                        jdn--;

                    // Convert JDN to Gregorian date using standard algorithm
                    int a = jdn + 32044;
                    int b = (4 * a + 3) / 146097;
                    int c = a - (146097 * b) / 4;
                    int d = (4 * c + 3) / 1461;
                    int e = c - (1461 * d) / 4;
                    int m = (5 * e + 2) / 153;

                    result.day = e - (153 * m + 2) / 5 + 1;
                    result.month = m + 3 - 12 * (m / 10);
                    result.year = 100 * b + d - 4800 + m / 10;
                }

                double integer_part;
                double fractional_part = std::modf(number, &integer_part);

                // Convert fractional part to total seconds
                double total_seconds = fractional_part * 24 * 60 * 60;
                
                // Handle negative fractional part (for negative numbers)
                if (total_seconds < 0) {
                    total_seconds += 86400.0;
                    integer_part -= 1;
                }
                
                // Handle day overflow
                while (total_seconds >= 86400.0) {
                    total_seconds -= 86400.0;
                    integer_part += 1;
                }
                
                // Extract hours, minutes, seconds, and microseconds
                result.hour = static_cast<int>(total_seconds / 3600);
                result.minute = static_cast<int>((total_seconds - result.hour * 3600) / 60);
                double seconds_with_fraction = total_seconds - result.hour * 3600 - result.minute * 60;
                result.second = static_cast<int>(seconds_with_fraction);
                // Use floor to avoid rounding errors that can propagate to minutes
                result.microsecond = static_cast<int>(std::floor((seconds_with_fraction - result.second) * 1000000));
                
                // Clamp microsecond to valid range
                if (result.microsecond < 0) result.microsecond = 0;
                if (result.microsecond >= 1000000) result.microsecond = 999999;

                return result;
            }

            double to_number(bool is_date1904 = false) const
            {
                int days_since_1900;
                if (day == 29 && month == 2 && year == 1900)
                {
                    days_since_1900 = 60;
                }
                else
                {
                    int m_adj = (month - 14) / 12;
                    days_since_1900 = (1461 * (year + 4800 + m_adj)) / 4
                                    + (367 * (month - 2 - 12 * m_adj)) / 12
                                    - (3 * ((year + 4900 + m_adj) / 100)) / 4
                                    + day - 2415020 - 32075;

                    if (days_since_1900 <= 60) days_since_1900--;
                    if (is_date1904) days_since_1900 -= 1462;
                }

                constexpr std::uint64_t us_per_hour = 3600000000ULL;
                std::uint64_t microseconds = microsecond + second * 1000000ULL + minute * 60000000ULL + hour * us_per_hour;
                double time_part = microseconds / (24.0 * us_per_hour);
                time_part = std::floor(time_part * 100000000000ULL + 0.5) / 100000000000ULL;

                return days_since_1900 + time_part;
            }
        };

        struct format_condition
        {
            enum class condition_type
            {
                less_than,
                less_or_equal,
                equal,
                not_equal,
                greater_than,
                greater_or_equal
            } type = condition_type::not_equal;

            double value = 0.0;

            bool satisfied_by(double number) const
            {
                // Ordered by likely frequency of use
                switch (type)
                {
                case condition_type::greater_than: return number > value + float_epsilon;
                case condition_type::less_than: return number < value - float_epsilon;
                case condition_type::greater_or_equal: return number >= value - float_epsilon;
                case condition_type::less_or_equal: return number <= value + float_epsilon;
                case condition_type::equal: return std::fabs(number - value) < float_epsilon;
                case condition_type::not_equal: return std::fabs(number - value) >= float_epsilon;
                }
                return false;
            }
        };
        struct format_placeholders
        {
            enum class placeholders_type
            {
                general,
                text,
                integer_only,
                integer_part,
                fractional_part,
                fraction_integer,
                fraction_numerator,
                fraction_denominator,
                scientific_significand,
                scientific_exponent_plus,
                scientific_exponent_minus
            } type = placeholders_type::general;

            bool use_comma_separator = false;
            bool percentage = false;
            bool scientific = false;

            std::size_t num_zeros = 0;     // 0
            std::size_t num_optionals = 0; // #
            std::size_t num_spaces = 0;    // ?
            std::size_t thousands_scale = 0;
        };
        struct number_format_token
        {
            enum class token_type
            {
                color,
                locale,
                condition,
                text,
                fill,
                space,
                number,
                datetime,
                end_section,
            };
            token_type type = token_type::end_section;

            std::string string;
        };
        struct template_part
        {
            enum class template_type
            {
                text,
                fill,
                space,
                general,
                month_number,
                month_number_leading_zero,
                month_abbreviation,
                month_name,
                month_letter,
                day_number,
                day_number_leading_zero,
                day_abbreviation,
                day_name,
                year_short,
                year_long,
                hour,
                hour_leading_zero,
                minute,
                minute_leading_zero,
                second,
                second_fractional,
                second_leading_zero,
                second_leading_zero_fractional,
                am_pm,
                a_p,
                elapsed_hours,
                elapsed_minutes,
                elapsed_seconds
            } type = template_type::general;

            std::string string;
            format_placeholders placeholders;
        };
        struct format_code
        {
            bool has_color = false;
            bool has_locale = false;
            bool is_datetime = false;
            bool is_timedelta = false;
            bool twelve_hour = false;
            bool has_condition = false;
            format_condition condition;
            std::vector<template_part> parts;
        };

    private:
        std::string _format_string;
        std::vector<format_code> _format;

    public:
        number_format(const std::string &format_string) : _format_string(format_string)
        {
            format_code section;
            template_part part;

            auto tokens = parse_tokens();
            for (const auto &token : tokens)
            {
                switch (token.type)
                {
                case number_format_token::token_type::end_section:
                    _format.push_back(section);
                    section = format_code();
                    break;
                case number_format_token::token_type::color:
                    if (section.has_color || section.has_condition || section.has_locale || !section.parts.empty())
                        throw std::runtime_error("color should be the first part of a format");
                    section.has_color = true;
                    break;

                case number_format_token::token_type::locale:
                {
                    if (section.has_locale)
                        throw std::runtime_error("multiple locales");

                    section.has_locale = true;

                    auto hyphen_index = token.string.find('-');
                    if (hyphen_index == std::string::npos || token.string.front() != '$')
                        throw std::runtime_error("bad locale: " + token.string);

                    if (hyphen_index > 1)
                    {
                        part.type = template_part::template_type::text;
                        part.string = token.string.substr(1, hyphen_index - 1);
                        section.parts.push_back(part);
                        part = template_part();
                    }
                    break;
                }
                case number_format_token::token_type::condition:
                {
                    if (section.has_condition)
                        throw std::runtime_error("multiple conditions");

                    section.has_condition = true;
                    const auto& s = token.string;
                    char op1 = s[0], op2 = s.size() > 1 ? s[1] : '\0';
                    std::size_t value_start = 1;

                    if (op1 == '<')
                    {
                        if (op2 == '=') { section.condition.type = format_condition::condition_type::less_or_equal; value_start = 2; }
                        else if (op2 == '>') { section.condition.type = format_condition::condition_type::not_equal; value_start = 2; }
                        else section.condition.type = format_condition::condition_type::less_than;
                    }
                    else if (op1 == '>')
                    {
                        if (op2 == '=') { section.condition.type = format_condition::condition_type::greater_or_equal; value_start = 2; }
                        else section.condition.type = format_condition::condition_type::greater_than;
                    }
                    else if (op1 == '=')
                        section.condition.type = format_condition::condition_type::equal;

                    section.condition.value = std::stod(s.substr(value_start));
                    break;
                }
                case number_format_token::token_type::text:
                case number_format_token::token_type::fill:
                case number_format_token::token_type::space:
                    part.type = (token.type == number_format_token::token_type::text) ? template_part::template_type::text
                                : (token.type == number_format_token::token_type::fill) ? template_part::template_type::fill
                                : template_part::template_type::space;
                    part.string = token.string;
                    section.parts.push_back(part);
                    part = template_part();
                    break;
                case number_format_token::token_type::number:
                    part.type = template_part::template_type::general;
                    part.placeholders = parse_placeholders(token.string);
                    section.parts.push_back(part);
                    part = template_part();
                    break;
                case number_format_token::token_type::datetime:
                {
                    section.is_datetime = true;
                    const auto& str = token.string;

                    if (str.front() == '[')
                    {
                        section.is_timedelta = true;
                        if (str == "[h]" || str == "[hh]") part.type = template_part::template_type::elapsed_hours;
                        else if (str == "[m]" || str == "[mm]") part.type = template_part::template_type::elapsed_minutes;
                        else if (str == "[s]" || str == "[ss]") part.type = template_part::template_type::elapsed_seconds;
                        else throw std::runtime_error("unhandled");
                    }
                    else if (str[0] == 'm')
                    {
                        if (section.is_timedelta)
                        {
                            if (str == "m") part.type = template_part::template_type::minute;
                            else if (str == "mm") part.type = template_part::template_type::minute_leading_zero;
                            else throw std::runtime_error("unhandled");
                        }
                        else
                        {
                            if (str == "m") part.type = template_part::template_type::month_number;
                            else if (str == "mm") part.type = template_part::template_type::month_number_leading_zero;
                            else if (str == "mmm") part.type = template_part::template_type::month_abbreviation;
                            else if (str == "mmmm") part.type = template_part::template_type::month_name;
                            else if (str == "mmmmm") part.type = template_part::template_type::month_letter;
                            else throw std::runtime_error("unhandled");
                        }
                    }
                    else if (str[0] == 'd')
                    {
                        if (str == "d") part.type = template_part::template_type::day_number;
                        else if (str == "dd") part.type = template_part::template_type::day_number_leading_zero;
                        else if (str == "ddd") part.type = template_part::template_type::day_abbreviation;
                        else if (str == "dddd") part.type = template_part::template_type::day_name;
                        else throw std::runtime_error("unhandled");
                    }
                    else if (str[0] == 'y')
                    {
                        if (str == "yy") part.type = template_part::template_type::year_short;
                        else if (str == "yyyy") part.type = template_part::template_type::year_long;
                        else throw std::runtime_error("unhandled");
                    }
                    else if (str[0] == 'h')
                    {
                        if (str == "h") part.type = template_part::template_type::hour;
                        else if (str == "hh") part.type = template_part::template_type::hour_leading_zero;
                        else throw std::runtime_error("unhandled");
                    }
                    else if (str[0] == 's')
                    {
                        if (str == "s") part.type = template_part::template_type::second;
                        else if (str == "ss") part.type = template_part::template_type::second_leading_zero;
                        else throw std::runtime_error("unhandled");
                    }
                    else if (str[0] == 'A')
                    {
                        section.twelve_hour = true;
                        if (str == "AM/PM") part.type = template_part::template_type::am_pm;
                        else if (str == "A/P") part.type = template_part::template_type::a_p;
                        else throw std::runtime_error("unhandled");
                    }
                    else
                    {
                        throw std::runtime_error("unhandled");
                    }

                    section.parts.push_back(part);
                    part = template_part();
                    break;
                }
                }
            }
            parse_finalize();
        }
        void parse_finalize()
        {
            for (auto &code : _format)
            {
                bool fix = false;
                bool leading_zero = false;
                std::size_t minutes_index = 0;

                bool integer_part = false;
                bool fractional_part = false;
                std::size_t integer_part_index = 0;

                bool percentage = false;

                bool exponent = false;

                bool fraction = false;
                std::size_t fraction_denominator_index = 0;
                std::size_t fraction_numerator_index = 0;

                bool seconds = false;
                bool fractional_seconds = false;
                std::size_t seconds_index = 0;

                for (std::size_t i = 0; i < code.parts.size(); ++i)
                {
                    const auto &part = code.parts[i];

                    if (i > 0 && i + 1 < code.parts.size() && part.type == template_part::template_type::text && part.string == "/" && code.parts[i - 1].placeholders.type == format_placeholders::placeholders_type::integer_part && code.parts[i + 1].placeholders.type == format_placeholders::placeholders_type::integer_part)
                    {
                        fraction = true;
                        fraction_numerator_index = i - 1;
                        fraction_denominator_index = i + 1;
                    }

                    if (part.placeholders.type == format_placeholders::placeholders_type::integer_part)
                    {
                        integer_part = true;
                        integer_part_index = i;
                    }
                    else if (part.placeholders.type == format_placeholders::placeholders_type::fractional_part)
                    {
                        fractional_part = true;
                    }
                    else if (part.placeholders.type == format_placeholders::placeholders_type::scientific_exponent_plus || part.placeholders.type == format_placeholders::placeholders_type::scientific_exponent_minus)
                    {
                        exponent = true;
                    }

                    if (part.placeholders.percentage)
                        percentage = true;

                    if (part.type == template_part::template_type::second || part.type == template_part::template_type::second_leading_zero)
                    {
                        seconds = true;
                        seconds_index = i;
                    }

                    if (seconds && part.placeholders.type == format_placeholders::placeholders_type::fractional_part)
                        fractional_seconds = true;

                    // Detect if 'm/mm' should be minutes instead of months
                    if ((part.type == template_part::template_type::month_number || part.type == template_part::template_type::month_number_leading_zero) && !fix)
                    {
                        bool is_minute = false;

                        // Check if followed by seconds
                        if (code.parts.size() > 1 && i < code.parts.size() - 2)
                        {
                            const auto &next = code.parts[i + 1];
                            const auto &after_next = code.parts[i + 2];

                            if (next.type == template_part::template_type::second || next.type == template_part::template_type::second_leading_zero ||
                                (next.type == template_part::template_type::text && next.string == ":" &&
                                 (after_next.type == template_part::template_type::second || after_next.type == template_part::template_type::second_leading_zero)))
                                is_minute = true;
                        }

                        // Check if preceded by hours
                        if (!is_minute && i > 1)
                        {
                            const auto &prev = code.parts[i - 1];
                            const auto &prev2 = code.parts[i - 2];

                            if (prev.type == template_part::template_type::text && prev.string == ":" &&
                                (prev2.type == template_part::template_type::hour || prev2.type == template_part::template_type::hour_leading_zero))
                                is_minute = true;
                        }

                        if (is_minute)
                        {
                            fix = true;
                            leading_zero = part.type == template_part::template_type::month_number_leading_zero;
                            minutes_index = i;
                        }
                    }
                }

                if (fix)
                    code.parts[minutes_index].type = leading_zero ? template_part::template_type::minute_leading_zero : template_part::template_type::minute;

                if (integer_part && !fractional_part)
                    code.parts[integer_part_index].placeholders.type = format_placeholders::placeholders_type::integer_only;

                if (integer_part && fractional_part && percentage)
                    code.parts[integer_part_index].placeholders.percentage = true;

                if (exponent)
                    for (auto& p : code.parts) p.placeholders.scientific = true;

                if (fraction)
                {
                    code.parts[fraction_numerator_index].placeholders.type = format_placeholders::placeholders_type::fraction_numerator;
                    code.parts[fraction_denominator_index].placeholders.type = format_placeholders::placeholders_type::fraction_denominator;

                    for (auto& p : code.parts)
                        if (p.placeholders.type == format_placeholders::placeholders_type::integer_part)
                            p.placeholders.type = format_placeholders::placeholders_type::fraction_integer;
                }

                if (fractional_seconds)
                    code.parts[seconds_index].type = (code.parts[seconds_index].type == template_part::template_type::second)
                        ? template_part::template_type::second_fractional : template_part::template_type::second_leading_zero_fractional;
            }

            if (_format.size() > 4)
                throw std::runtime_error("too many format codes");

            if (_format.size() > 2 && _format[0].has_condition && _format[1].has_condition && _format[2].has_condition)
                throw std::runtime_error("format should have a maximum of two codes with conditions");
        }

        std::vector<number_format_token> parse_tokens() const
        {
            std::vector<number_format_token> tokens;
            std::string::size_type position = 0;
            while (position < _format_string.size())
            {
                number_format_token token;
                auto current_char = _format_string[position++];
                switch (current_char)
                {
                case '[':
                    if (position == _format_string.size())
                        throw std::runtime_error("missing ]");

                    if (_format_string[position] == ']')
                        throw std::runtime_error("empty []");

                    do
                    {
                        token.string.push_back(_format_string[position++]);
                    } while (position < _format_string.size() && _format_string[position] != ']');

                    if (token.string[0] == '<' || token.string[0] == '>' || token.string[0] == '=')
                        token.type = number_format_token::token_type::condition;
                    else if (token.string[0] == '$')
                        token.type = number_format_token::token_type::locale;
                    else if (token.string.size() <= 2 && (token.string == "h" || token.string == "hh" || token.string == "m" || token.string == "mm" || token.string == "s" || token.string == "ss"))
                    {
                        token.type = number_format_token::token_type::datetime;
                        token.string = "[" + token.string + "]";
                    }
                    else
                        token.type = number_format_token::token_type::color;

                    ++position;
                    break;

                case '\\':
                    token.type = number_format_token::token_type::text;
                    token.string.push_back(_format_string[position++]);
                    break;

                case 'G':
                    if (_format_string.substr(position - 1, 7) != "General")
                        throw std::runtime_error("expected General");
                    token.type = number_format_token::token_type::number;
                    token.string = "General";
                    position += 6;
                    break;

                case '_':
                    token.type = number_format_token::token_type::space;
                    token.string.push_back(_format_string[position++]);
                    break;

                case '*':
                    token.type = number_format_token::token_type::fill;
                    token.string.push_back(_format_string[position++]);
                    break;

                case '0':
                case '#':
                case '?':
                case '.':
                    token.type = number_format_token::token_type::number;
                    do
                    {
                        token.string.push_back(current_char);
                        current_char = _format_string[position++];
                    } while (current_char == '0' || current_char == '#' || current_char == '?' || current_char == ',');
                    --position;
                    if (current_char == '%')
                    {
                        token.string.push_back('%');
                        ++position;
                    }
                    break;

                case 'y':
                case 'Y':
                case 'm':
                case 'M':
                case 'd':
                case 'D':
                case 'h':
                case 'H':
                case 's':
                case 'S':
                    token.type = number_format_token::token_type::datetime;
                    token.string.push_back(static_cast<char>(std::tolower(static_cast<std::uint8_t>(current_char))));
                    while (_format_string[position] == current_char)
                    {
                        token.string.push_back(static_cast<char>(std::tolower(static_cast<std::uint8_t>(current_char))));
                        ++position;
                    }
                    break;

                case 'A':
                    token.type = number_format_token::token_type::datetime;
                    if (_format_string.substr(position - 1, 5) == "AM/PM")
                    {
                        position += 4;
                        token.string = "AM/PM";
                    }
                    else if (_format_string.substr(position - 1, 3) == "A/P")
                    {
                        position += 2;
                        token.string = "A/P";
                    }
                    else
                        throw std::runtime_error("expected AM/PM or A/P");
                    break;

                case '"':
                {
                    token.type = number_format_token::token_type::text;
                    auto start = position;
                    auto end = _format_string.find('"', position);

                    while (end != std::string::npos && _format_string[end - 1] == '\\')
                    {
                        token.string.append(_format_string.substr(start, end - start - 1));
                        token.string.push_back('"');
                        position = end + 1;
                        start = position;
                        end = _format_string.find('"', position);
                    }

                    if (end != start)
                        token.string.append(_format_string.substr(start, end - start));

                    position = end + 1;
                    break;
                }

                case ';':
                    token.type = number_format_token::token_type::end_section;
                    break;

                case '(':
                case ')':
                case '-':
                case '+':
                case ':':
                case ' ':
                case '/':
                    token.type = number_format_token::token_type::text;
                    token.string.push_back(current_char);
                    break;

                case '@':
                    token.type = number_format_token::token_type::number;
                    token.string.push_back(current_char);
                    break;

                case 'E':
                    token.type = number_format_token::token_type::number;
                    token.string.push_back(current_char);
                    current_char = _format_string[position++];
                    if (current_char == '+' || current_char == '-')
                    {
                        token.string.push_back(current_char);
                        current_char = _format_string[position++];
                        while (current_char == '0' || current_char == '#' || current_char == '?')
                        {
                            token.string.push_back(current_char);
                            if (position < _format_string.size())
                                current_char = _format_string[position++];
                            else
                                break;
                        }
                        if (position <= _format_string.size())
                            --position;
                    }
                    break;

                default:
                    throw std::runtime_error("unexpected character");
                }
                tokens.push_back(token);
            }
            
            if (tokens.empty() || tokens.back().type != number_format_token::token_type::end_section)
                tokens.push_back({number_format_token::token_type::end_section, ""});

            return tokens;
        }
        format_placeholders parse_placeholders(const std::string &s) const
        {
            format_placeholders p;

            if (s == "General") { p.type = format_placeholders::placeholders_type::general; return p; }
            if (s == "@") { p.type = format_placeholders::placeholders_type::text; return p; }
            
            char first = s.front();
            if (first == '.') p.type = format_placeholders::placeholders_type::fractional_part;
            else if (first == 'E')
            {
                p.type = (s[1] == '+') ? format_placeholders::placeholders_type::scientific_exponent_plus 
                                        : format_placeholders::placeholders_type::scientific_exponent_minus;
                for (std::size_t i = 2; i < s.size(); ++i)
                    if (s[i] == '0') ++p.num_zeros;
                    else if (s[i] == '#') ++p.num_optionals;
                return p;
            }
            else p.type = format_placeholders::placeholders_type::integer_part;

            if (s.back() == '%') p.percentage = true;

            // Count trailing commas for thousands_scale
            std::size_t len = s.size();
            while (len > 0 && s[len - 1] == ',') { ++p.thousands_scale; --len; }
            p.use_comma_separator = (s.find(',') != std::string::npos && p.thousands_scale < s.size());

            // Count placeholders (excluding trailing commas)
            for (std::size_t i = 0; i < len; ++i)
            {
                char c = s[i];
                if (c == '0') ++p.num_zeros;
                else if (c == '#') ++p.num_optionals;
                else if (c == '?') ++p.num_spaces;
            }

            return p;
        }

        std::string format(double number, bool is_date1904 = false)
        {
            // Handle special floating-point values
            if (std::isnan(number))
                return "#NUM!";
            if (std::isinf(number))
                return number > 0 ? "#INF!" : "-#INF!";

            // With conditions
            if (_format[0].has_condition)
            {
                if (_format[0].condition.satisfied_by(number))
                    return format_number(_format[0], number, is_date1904);
                if (_format.size() == 1)
                    return std::string(default_cell_width, '#');
                if (!_format[1].has_condition || _format[1].condition.satisfied_by(number))
                    return format_number(_format[1], number, is_date1904);
                if (_format.size() == 2)
                    return std::string(default_cell_width, '#');
                return format_number(_format[2], number, is_date1904);
            }

            // No conditions - format based on sign
            switch (_format.size())
            {
            case 1: return format_number(_format[0], number, is_date1904);
            case 2: return format_number(_format[number >= 0 ? 0 : 1], number, is_date1904);
            default:
                if (number > 0) return format_number(_format[0], number, is_date1904);
                if (number < 0) return format_number(_format[1], number, is_date1904);
                return format_number(_format[2], number, is_date1904);
            }
        }
        std::string format(const std::string &text)
        {
            // If 4 sections, use 4th for text; otherwise use 1st section
            std::size_t format_index = (_format.size() >= 4) ? 3 : 0;
            return format_text(_format[format_index], text);
        }

        // Non-throwing format methods
        bool try_format(double number, std::string& result, bool is_date1904 = false) noexcept
        {
            try
            {
                result = format(number, is_date1904);
                return true;
            }
            catch (...)
            {
                result.clear();
                return false;
            }
        }

        bool try_format(const std::string& text, std::string& result) noexcept
        {
            try
            {
                result = format(text);
                return true;
            }
            catch (...)
            {
                result.clear();
                return false;
            }
        }

    private:
        std::string fill_placeholders(const format_placeholders &p, double number)
        {
            std::string result;

            // Handle General format
            if (p.type == format_placeholders::placeholders_type::general || p.type == format_placeholders::placeholders_type::text)
            {
                if (std::fabs(number) < float_epsilon) return "0";
                
                double abs_num = std::fabs(number);
                
                // Scientific notation for very small/large numbers
                if (abs_num < general_small_threshold || abs_num >= general_large_threshold)
                {
                    int exponent = static_cast<int>(std::floor(std::log10(abs_num)));
                    double mantissa = number / std::pow(10.0, exponent);
                    
                    char buffer[32];
                    std::snprintf(buffer, sizeof(buffer), "%.15g", mantissa);
                    
                    std::string exp_str = std::to_string(std::abs(exponent));
                    if (exp_str.size() == 1) exp_str = "0" + exp_str;
                    
                    return std::string(buffer) + "e" + (exponent < 0 ? "-" : "+") + exp_str;
                }
                
                char buffer[64];
                std::snprintf(buffer, sizeof(buffer), "%.15g", number);
                return std::string(buffer);
            }

            // Apply transformations
            if (p.percentage) number *= 100;
            if (p.thousands_scale > 0) number /= std::pow(1000.0, p.thousands_scale);
            if (p.type == format_placeholders::placeholders_type::integer_only) number = std::round(number);

            auto int_val = static_cast<long long>(number);
            double frac_val = number - static_cast<double>(int_val);
            bool has_frac = std::fabs(frac_val) > float_epsilon;

            // Integer part types
            if (p.type == format_placeholders::placeholders_type::integer_only ||
                p.type == format_placeholders::placeholders_type::integer_part ||
                p.type == format_placeholders::placeholders_type::fraction_integer)
            {
                result = (int_val == 0 && p.num_zeros == 0) ? (has_frac ? "0" : "") : std::to_string(int_val);

                // Pad with zeros
                if (result.size() < p.num_zeros)
                    result.insert(0, p.num_zeros - result.size(), '0');

                // Pad with spaces
                if (result.size() < p.num_zeros + p.num_spaces)
                    result.insert(0, p.num_zeros + p.num_spaces - result.size(), ' ');

                // Add comma separators
                if (p.use_comma_separator && !result.empty())
                {
                    std::string temp;
                    for (std::size_t i = 0; i < result.size(); ++i)
                    {
                        temp.push_back(result[i]);
                        if ((result.size() - i - 1) > 0 && (result.size() - i - 1) % 3 == 0)
                            temp.push_back(',');
                    }
                    result = std::move(temp);
                }

                if (p.percentage && p.type == format_placeholders::placeholders_type::integer_only)
                    result.push_back('%');
            }
            // Fractional part
            else if (p.type == format_placeholders::placeholders_type::fractional_part)
            {
                std::size_t decimal_places = p.num_zeros + p.num_optionals + p.num_spaces;
                double scale = std::pow(10.0, decimal_places);
                double frac_part = std::round(number * scale) / scale;
                frac_part -= static_cast<long long>(frac_part);

                if (std::fabs(frac_part) < 1e-15) frac_part = 0.0;

                if (frac_part == 0.0 && p.num_zeros == 0)
                {
                    result = "";
                }
                else
                {
                    result = (frac_part == 0.0) ? "." : [&]() {
                        char buffer[64];
                        std::snprintf(buffer, sizeof(buffer), "%.*f", static_cast<int>(decimal_places), std::fabs(frac_part));
                        std::string full_str(buffer);
                        auto dot_pos = full_str.find('.');
                        return (dot_pos != std::string::npos) ? full_str.substr(dot_pos) : ".";
                    }();

                    std::size_t total_digits = p.num_zeros + p.num_optionals + p.num_spaces;

                    // Trim to required precision
                    while (result.size() > total_digits + 1) result.pop_back();

                    // Remove trailing zeros for optional placeholders
                    std::size_t min_len = p.num_zeros + 1;
                    while (result.size() > min_len && result.back() == '0') result.pop_back();

                    if (result == "." && p.num_zeros == 0) result = "";

                    // Pad zeros and spaces
                    while (result.size() > 0 && result.size() < p.num_zeros + 1) result.push_back('0');
                    while (p.num_spaces > 0 && result.size() < total_digits + 1) result.push_back(' ');
                }

                if (p.percentage) result.push_back('%');
            }

            return result;
        }
        std::string fill_fraction_placeholders(const format_placeholders &numerator, const format_placeholders &denominator, double number, bool improper)
        {
            // Get the fractional part
            double int_part;
            double fractional_part = std::modf(number, &int_part);
            if (fractional_part < 0) fractional_part = -fractional_part;
            
            // Handle zero fractional part
            std::size_t denom_digits = denominator.num_zeros + denominator.num_optionals + denominator.num_spaces;
            if (denom_digits == 0) denom_digits = 1;
            int default_denom = static_cast<int>(std::pow(10.0, denom_digits));
            
            if (fractional_part < 1e-10)
                return "0/" + std::to_string(default_denom);
            
            int max_denominator = default_denom - 1;
            if (max_denominator < 1) max_denominator = 99;
            
            // Find best fraction using continued fraction algorithm
            int best_num = 0, best_den = 1;
            double best_error = fractional_part;
            
            // Try common denominators first
            static constexpr int common_denoms[] = {2, 4, 8, 16, 32, 64, 10, 100, 3, 6, 12, 24, 48, 96, 5, 20, 50};
            for (int d : common_denoms)
            {
                if (d > max_denominator) continue;
                int n = static_cast<int>(std::round(fractional_part * d));
                double error = std::fabs(fractional_part - static_cast<double>(n) / d);
                if (error < best_error) { best_error = error; best_num = n; best_den = d; }
            }
            
            // Only try all denominators if common ones didn't give good enough result
            if (best_error > 1e-6)
            {
                for (int d = 1; d <= max_denominator; ++d)
                {
                    int n = static_cast<int>(std::round(fractional_part * d));
                    double error = std::fabs(fractional_part - static_cast<double>(n) / d);
                    if (error < best_error - 1e-10) { best_error = error; best_num = n; best_den = d; }
                }
            }

            int g = compute_gcd(best_num, best_den);
            return std::to_string(best_num / g) + "/" + std::to_string(best_den / g);
        }
        std::string fill_scientific_placeholders(const format_placeholders &integer_part, const format_placeholders &fractional_part, const format_placeholders &exponent_part, double number)
        {
            std::string result;

            if (number == 0.0)
            {
                result = std::string(integer_part.num_zeros + integer_part.num_optionals, '0');
                result.push_back('.');
                result.append(std::string(fractional_part.num_zeros, '0'));

                // Remove trailing zeros for optional placeholders
                for (std::size_t i = 0; i < fractional_part.num_optionals && result.back() == '0'; ++i)
                    result.pop_back();

                if (result.back() == '.') result.pop_back();

                result.append(exponent_part.type == format_placeholders::placeholders_type::scientific_exponent_plus ? "E+00" : "E00");
                return result;
            }

            // Calculate exponent (can be negative for numbers < 1)
            int exponent = static_cast<int>(std::floor(std::log10(std::fabs(number))));

            // Normalize: number = mantissa * 10^exponent, where 1 <= mantissa < 10
            // Use std::pow directly since exponent can be negative
            double mantissa = number / std::pow(10.0, exponent);

            // Round mantissa based on fractional placeholders
            std::size_t decimal_places = fractional_part.num_zeros + fractional_part.num_optionals;
            if (decimal_places > 0)
            {
                double scale = std::pow(10.0, decimal_places);
                mantissa = std::round(mantissa * scale) / scale;
                // Check if rounding pushed mantissa to >= 10
                if (std::fabs(mantissa) >= 10.0)
                {
                    mantissa /= 10.0;
                    exponent++;
                }
            }

            // Format mantissa
            bool is_negative = mantissa < 0;
            mantissa = std::fabs(mantissa);

            int int_part = static_cast<int>(mantissa);
            double frac_part = mantissa - int_part;

            // Integer part of mantissa (should be single digit for standard scientific notation)
            result = std::to_string(int_part);

            // Pad integer part only for required (0) placeholders, not for optional (#) placeholders
            while (result.size() < integer_part.num_zeros)
                result.insert(0, "0");

            // Fractional part
            if (fractional_part.num_zeros + fractional_part.num_optionals > 0)
            {
                std::string frac_str;
                if (frac_part > 0 || fractional_part.num_zeros > 0)
                {
                    frac_str = std::to_string(frac_part).substr(2); // Skip "0."
                    std::size_t max_len = fractional_part.num_zeros + fractional_part.num_optionals;
                    if (frac_str.size() > max_len)
                        frac_str = frac_str.substr(0, max_len);
                }
                while (frac_str.size() < fractional_part.num_zeros)
                    frac_str.push_back('0');
                while (frac_str.size() < fractional_part.num_zeros + fractional_part.num_spaces)
                    frac_str.push_back(' ');

                if (!frac_str.empty())
                {
                    result.push_back('.');
                    result.append(frac_str);
                }
            }

            // Exponent part
            bool exp_negative = exponent < 0;
            std::string exp_str = std::to_string(std::abs(exponent));

            std::size_t exp_digits = exponent_part.num_zeros + exponent_part.num_optionals;
            while (exp_str.size() < exp_digits)
                exp_str.insert(0, "0");

            if (exponent_part.type == format_placeholders::placeholders_type::scientific_exponent_plus)
                result.append(exp_negative ? "E-" : "E+");
            else
            {
                result.push_back('E');
                if (exp_negative) result.push_back('-');
            }
            result.append(exp_str);

            if (is_negative)
                result.insert(0, "-");

            return result;
        }
        static constexpr std::array<const char*, 12> month_names = {
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        };
        static constexpr std::array<const char*, 7> day_names = {
            "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
        };

        // Helper to append integer with optional leading zero
        static void append_with_leading_zero(std::string& result, int value)
        {
            if (value < 10) result.push_back('0');
            result.append(std::to_string(value));
        }

        // Helper to get timedelta minutes
        static int get_timedelta_minutes(double number)
        {
            return static_cast<long long>(number * 24 * 60) % 60;
        }

        // Helper to compute GCD
        static int compute_gcd(int a, int b)
        {
            while (b) { int t = b; b = a % b; a = t; }
            return a;
        }

        // Helper to check if format has explicit negative handling
        static bool has_explicit_negative_format(const format_code& format)
        {
            for (const auto& part : format.parts)
            {
                if (part.type == template_part::template_type::text &&
                    (part.string == "(" || part.string == ")" || part.string == "-"))
                    return true;
            }
            return false;
        }

        // Helper to check if format is scientific
        static bool is_scientific_format(const format_code& format)
        {
            for (const auto& part : format.parts)
            {
                if (part.placeholders.scientific)
                    return true;
            }
            return false;
        }

        std::string format_number(const format_code &format, double number, bool is_date1904 = false)
        {

            std::string result;
            result.reserve(32); // Pre-allocate for typical format output

            bool is_negative = number < 0;
            
            if (is_negative && format.is_datetime)
            {
                return std::string(default_cell_width, '#');
            }

            number = std::fabs(number);

            bool has_explicit_negative = has_explicit_negative_format(format);

            // Pre-round number for decimal formats to handle carry correctly
            bool is_scientific = is_scientific_format(format);
            
            if (!format.is_datetime && !format.is_timedelta && !is_scientific)
            {
                std::size_t decimal_places = 0;
                bool has_fractional = false;
                bool is_percentage = false;
                std::size_t thousands_scale = 0;

                for (const auto &part : format.parts)
                {
                    if (part.placeholders.type == format_placeholders::placeholders_type::fractional_part)
                    {
                        has_fractional = true;
                        decimal_places = part.placeholders.num_zeros + part.placeholders.num_optionals + part.placeholders.num_spaces;
                    }
                    if (part.placeholders.percentage)
                    {
                        is_percentage = true;
                    }
                    if (part.placeholders.thousands_scale > thousands_scale)
                    {
                        thousands_scale = part.placeholders.thousands_scale;
                    }
                }

                if (has_fractional && decimal_places > 0)
                {
                    double n = number;
                    if (is_percentage) n *= 100;
                    if (thousands_scale > 0) n /= std::pow(1000.0, thousands_scale);
                    double scale = std::pow(10.0, decimal_places);
                    number = std::round(n * scale) / scale;
                    if (is_percentage) number /= 100;
                    if (thousands_scale > 0) number *= std::pow(1000.0, thousands_scale);
                }
            }

            datetime dt(0, 1, 0);
            std::size_t hour = 0;

            if (format.is_datetime)
            {
                if (number != 0.0)
                    dt = datetime::from_number(number, is_date1904);

                hour = static_cast<std::size_t>(dt.hour);

                if (format.twelve_hour)
                {
                    hour %= 12;

                    if (hour == 0)
                    {
                        hour = 12;
                    }
                }
            }

            bool improper_fraction = true;
            std::size_t fill_index = 0;
            bool fill = false;
            std::string fill_character;

            for (std::size_t i = 0; i < format.parts.size(); ++i)
            {
                const auto &part = format.parts[i];

                switch (part.type)
                {
                case template_part::template_type::space:
                    result.push_back(' ');
                    break;

                case template_part::template_type::text:
                    result.append(part.string);
                    break;

                case template_part::template_type::fill:
                    fill = true;
                    fill_index = result.size();
                    fill_character = part.string;
                    break;

                case template_part::template_type::general:
                {
                    if (part.placeholders.type == format_placeholders::placeholders_type::fractional_part && (format.is_datetime || format.is_timedelta))
                    {
                        auto digits = std::min(static_cast<std::size_t>(6), part.placeholders.num_zeros + part.placeholders.num_optionals);
                        auto denominator = static_cast<int>(std::pow(10.0, digits));
                        auto fractional_seconds = dt.microsecond / 1.0E6 * denominator;
                        fractional_seconds = std::round(fractional_seconds) / denominator;
                        result.append(fill_placeholders(part.placeholders, fractional_seconds));
                        break;
                    }

                    if (part.placeholders.type == format_placeholders::placeholders_type::fraction_integer)
                    {
                        improper_fraction = false;
                    }

                    if (part.placeholders.type == format_placeholders::placeholders_type::fraction_numerator)
                    {
                        i += 2;

                        if (number == 0.0)
                        {
                            result.pop_back();
                            break;
                        }

                        result.append(fill_fraction_placeholders(part.placeholders, format.parts[i].placeholders, number, improper_fraction));
                    }
                    else if (part.placeholders.scientific && part.placeholders.type == format_placeholders::placeholders_type::integer_part)
                    {
                        auto integer_part = part.placeholders;
                        ++i;
                        auto fractional_part = format.parts[i++].placeholders;
                        auto exponent_part = format.parts[i++].placeholders;
                        result.append(fill_scientific_placeholders(
                            integer_part, fractional_part, exponent_part, number));
                    }
                    else
                    {
                        result.append(fill_placeholders(part.placeholders, number));
                    }

                    break;
                }

                case template_part::template_type::day_number:
                    result.append(std::to_string(dt.day));
                    break;

                case template_part::template_type::day_number_leading_zero:
                    append_with_leading_zero(result, dt.day);
                    break;

                case template_part::template_type::month_abbreviation:
                    result.append(month_names[static_cast<std::size_t>(dt.month) - 1], 3);
                    break;

                case template_part::template_type::month_name:
                    result.append(month_names[static_cast<std::size_t>(dt.month) - 1]);
                    break;

                case template_part::template_type::month_number:
                    result.append(std::to_string(dt.month));
                    break;

                case template_part::template_type::month_number_leading_zero:
                    append_with_leading_zero(result, dt.month);
                    break;

                case template_part::template_type::year_short:
                    append_with_leading_zero(result, dt.year % 100);
                    break;

                case template_part::template_type::year_long:
                    result.append(std::to_string(dt.year));
                    break;

                case template_part::template_type::hour:
                    result.append(std::to_string(hour));
                    break;

                case template_part::template_type::hour_leading_zero:
                    append_with_leading_zero(result, static_cast<int>(hour));
                    break;

                case template_part::template_type::minute:
                {
                    int mins = format.is_timedelta ? get_timedelta_minutes(number) : dt.minute;
                    result.append(std::to_string(mins));
                    break;
                }

                case template_part::template_type::minute_leading_zero:
                {
                    int mins = format.is_timedelta ? get_timedelta_minutes(number) : dt.minute;
                    append_with_leading_zero(result, mins);
                    break;
                }

                case template_part::template_type::second:
                case template_part::template_type::second_fractional:
                    result.append(std::to_string(dt.second));
                    break;

                case template_part::template_type::second_leading_zero:
                case template_part::template_type::second_leading_zero_fractional:
                    append_with_leading_zero(result, dt.second);
                    break;

                case template_part::template_type::am_pm:
                    result.append(dt.hour < 12 ? "AM" : "PM");
                    break;

                case template_part::template_type::a_p:
                    result.append(dt.hour < 12 ? "A" : "P");
                    break;

                case template_part::template_type::elapsed_hours:
                    result.append(std::to_string(static_cast<long long>(number * 24)));
                    break;

                case template_part::template_type::elapsed_minutes:
                    result.append(std::to_string(static_cast<long long>(number * 1440)));
                    break;

                case template_part::template_type::elapsed_seconds:
                    result.append(std::to_string(static_cast<long long>(number * 86400)));
                    break;

                case template_part::template_type::month_letter:
                    result.push_back(month_names[static_cast<std::size_t>(dt.month) - 1][0]);
                    break;

                case template_part::template_type::day_abbreviation:
                    result.append(day_names[static_cast<std::size_t>(dt.weekday())], 3);
                    break;

                case template_part::template_type::day_name:
                    result.append(day_names[static_cast<std::size_t>(dt.weekday())]);
                    break;
                }
            }

            if (fill && result.size() < default_cell_width)
            {
                std::string fill_string(default_cell_width - result.size(), fill_character.front());
                result = result.substr(0, fill_index) + fill_string + result.substr(fill_index);
            }

            // Add negative sign if needed and format doesn't have explicit negative handling
            if (is_negative && !has_explicit_negative)
            {
                result.insert(0, "-");
            }

            return result;
        }
        std::string format_text(const format_code &format, const std::string &text)
        {
            std::string result;
            bool has_text_placeholder = false;

            for (const auto &part : format.parts)
            {
                if (part.type == template_part::template_type::text)
                    result.append(part.string);
                else if (part.type == template_part::template_type::general &&
                    (part.placeholders.type == format_placeholders::placeholders_type::text ||
                     part.placeholders.type == format_placeholders::placeholders_type::general))
                {
                    result.append(text);
                    has_text_placeholder = true;
                }
            }

            // If no text placeholder but has other parts, return original text
            if (!has_text_placeholder && !format.parts.empty())
            {
                for (const auto& p : format.parts)
                {
                    if (p.type != template_part::template_type::text &&
                        p.type != template_part::template_type::space &&
                        p.type != template_part::template_type::fill)
                        return text;
                }
            }

            return result;
        }
    };
}