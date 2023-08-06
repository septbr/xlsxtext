#pragma once

#include "miniz/miniz.h"
#include "pugixml/pugixml.hpp"

#include <cstring>
#include <string>
#include <vector>
#include <map>
#include <memory>

namespace xlsxtext
{
    class number_format
    {
    private:
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
                datetime result(0, 0, static_cast<int>(number));

                if (is_date1904)
                    result.day += 1462;

                if (result.day == 60)
                {
                    result.day = 29;
                    result.month = 2;
                    result.year = 1900;
                }
                else
                {
                    if (result.day < 60)
                        result.day++;

                    int l = result.day + 68569 + 2415019;
                    int n = int((4 * l) / 146097);
                    l = l - int((146097 * n + 3) / 4);
                    int i = int((4000 * (l + 1)) / 1461001);
                    l = l - int((1461 * i) / 4) + 31;
                    int j = int((80 * l) / 2447);
                    result.day = l - int((2447 * j) / 80);
                    l = int(j / 11);
                    result.month = j + 2 - (12 * l);
                    result.year = 100 * (n - 49) + i + l;
                }

                double integer_part;
                double fractional_part = std::modf(number, &integer_part);

                fractional_part *= 24;
                result.hour = static_cast<int>(fractional_part);
                fractional_part = 60 * (fractional_part - result.hour);
                result.minute = static_cast<int>(fractional_part);
                fractional_part = 60 * (fractional_part - result.minute);
                result.second = static_cast<int>(fractional_part);
                fractional_part = 1000000 * (fractional_part - result.second);
                result.microsecond = static_cast<int>(fractional_part);

                if (result.microsecond == 999999 && fractional_part - result.microsecond > 0.5)
                {
                    result.microsecond = 0;
                    result.second += 1;

                    if (result.second == 60)
                    {
                        result.second = 0;
                        result.minute += 1;

                        // TODO: too much nesting
                        if (result.minute == 60)
                        {
                            result.minute = 0;
                            result.hour += 1;
                        }
                    }
                }

                return result;
            }

            double to_number(bool is_date1904 = false) const
            {
                int days_since_1900 = 0;
                if (day == 29 && month == 2 && year == 1900)
                    days_since_1900 = 60;
                else
                {
                    days_since_1900 = int((1461 * (year + 4800 + int((month - 14) / 12))) / 4) + int((367 * (month - 2 - 12 * ((month - 14) / 12))) / 12) - int((3 * (int((year + 4900 + int((month - 14) / 12)) / 100))) / 4) + day - 2415019 - 32075;

                    if (days_since_1900 <= 60)
                        days_since_1900--;

                    if (is_date1904)
                        days_since_1900 -= 1462;
                }

                std::uint64_t microseconds = static_cast<std::uint64_t>(microsecond);
                microseconds += static_cast<std::uint64_t>(second * 1e6);
                microseconds += static_cast<std::uint64_t>(minute * 1e6 * 60);
                auto microseconds_per_hour = static_cast<std::uint64_t>(1e6) * 60 * 60;
                microseconds += static_cast<std::uint64_t>(hour) * microseconds_per_hour;
                auto number = microseconds / (24.0 * microseconds_per_hour);
                auto hundred_billion = static_cast<std::uint64_t>(1e9) * 100;
                number = std::floor(number * hundred_billion + 0.5) / hundred_billion;

                return days_since_1900 + number;
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
                switch (type)
                {
                case condition_type::greater_or_equal:
                    return number >= value;
                case condition_type::greater_than:
                    return number > value;
                case condition_type::less_or_equal:
                    return number <= value;
                case condition_type::less_than:
                    return number < value;
                case condition_type::not_equal:
                    return std::fabs(number - value) != 0.0;
                case condition_type::equal:
                    return std::fabs(number - value) == 0.0;
                }

                throw std::runtime_error("unhandled_switch_case");
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
            } type = token_type::end_section;

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
                {
                    _format.push_back(section);
                    section = format_code();
                    break;
                }
                case number_format_token::token_type::color:
                {
                    if (section.has_color || section.has_condition || section.has_locale || !section.parts.empty())
                    {
                        throw std::runtime_error(
                            "color should be the first part of a format");
                    }

                    section.has_color = true;
                    break;
                }
                case number_format_token::token_type::locale:
                {
                    if (section.has_locale)
                    {
                        throw std::runtime_error("multiple locales");
                    }

                    section.has_locale = true;

                    auto hyphen_index = token.string.find('-');
                    if (hyphen_index == std::string::npos || token.string.front() != '$')
                    {
                        throw std::runtime_error("bad locale: " + token.string);
                    }
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
                    {
                        throw std::runtime_error("multiple conditions");
                    }

                    section.has_condition = true;
                    std::string value;

                    if (token.string.front() == '<')
                    {
                        if (token.string[1] == '=')
                        {
                            section.condition.type = format_condition::condition_type::less_or_equal;
                            value = token.string.substr(2);
                        }
                        else if (token.string[1] == '>')
                        {
                            section.condition.type = format_condition::condition_type::not_equal;
                            value = token.string.substr(2);
                        }
                        else
                        {
                            section.condition.type = format_condition::condition_type::less_than;
                            value = token.string.substr(1);
                        }
                    }
                    else if (token.string.front() == '>')
                    {
                        if (token.string[1] == '=')
                        {
                            section.condition.type = format_condition::condition_type::greater_or_equal;
                            value = token.string.substr(2);
                        }
                        else
                        {
                            section.condition.type = format_condition::condition_type::greater_than;
                            value = token.string.substr(1);
                        }
                    }
                    else if (token.string.front() == '=')
                    {
                        section.condition.type = format_condition::condition_type::equal;
                        value = token.string.substr(1);
                    }

                    section.condition.value = std::stod(value);
                    break;
                }
                case number_format_token::token_type::text:
                {
                    part.type = template_part::template_type::text;
                    part.string = token.string;
                    section.parts.push_back(part);
                    part = template_part();

                    break;
                }
                case number_format_token::token_type::fill:
                {
                    part.type = template_part::template_type::fill;
                    part.string = token.string;
                    section.parts.push_back(part);
                    part = template_part();

                    break;
                }
                case number_format_token::token_type::space:
                {
                    part.type = template_part::template_type::space;
                    part.string = token.string;
                    section.parts.push_back(part);
                    part = template_part();

                    break;
                }
                case number_format_token::token_type::number:
                {
                    part.type = template_part::template_type::general;
                    part.placeholders = parse_placeholders(token.string);
                    section.parts.push_back(part);
                    part = template_part();

                    break;
                }
                case number_format_token::token_type::datetime:
                {
                    section.is_datetime = true;

                    switch (token.string.front())
                    {
                    case '[':
                        section.is_timedelta = true;

                        if (token.string == "[h]" || token.string == "[hh]")
                        {
                            part.type = template_part::template_type::elapsed_hours;
                            break;
                        }
                        else if (token.string == "[m]" || token.string == "[mm]")
                        {
                            part.type = template_part::template_type::elapsed_minutes;
                            break;
                        }
                        else if (token.string == "[s]" || token.string == "[ss]")
                        {
                            part.type = template_part::template_type::elapsed_seconds;
                            break;
                        }

                        throw std::runtime_error("unhandled");
                        break;

                    case 'm':
                        if (token.string == "m")
                        {
                            part.type = template_part::template_type::month_number;
                            break;
                        }
                        else if (token.string == "mm")
                        {
                            part.type = template_part::template_type::month_number_leading_zero;
                            break;
                        }
                        else if (token.string == "mmm")
                        {
                            part.type = template_part::template_type::month_abbreviation;
                            break;
                        }
                        else if (token.string == "mmmm")
                        {
                            part.type = template_part::template_type::month_name;
                            break;
                        }
                        else if (token.string == "mmmmm")
                        {
                            part.type = template_part::template_type::month_letter;
                            break;
                        }

                        throw std::runtime_error("unhandled");
                        break;

                    case 'd':
                        if (token.string == "d")
                        {
                            part.type = template_part::template_type::day_number;
                            break;
                        }
                        else if (token.string == "dd")
                        {
                            part.type = template_part::template_type::day_number_leading_zero;
                            break;
                        }
                        else if (token.string == "ddd")
                        {
                            part.type = template_part::template_type::day_abbreviation;
                            break;
                        }
                        else if (token.string == "dddd")
                        {
                            part.type = template_part::template_type::day_name;
                            break;
                        }

                        throw std::runtime_error("unhandled");
                        break;

                    case 'y':
                        if (token.string == "yy")
                        {
                            part.type = template_part::template_type::year_short;
                            break;
                        }
                        else if (token.string == "yyyy")
                        {
                            part.type = template_part::template_type::year_long;
                            break;
                        }

                        throw std::runtime_error("unhandled");
                        break;

                    case 'h':
                        if (token.string == "h")
                        {
                            part.type = template_part::template_type::hour;
                            break;
                        }
                        else if (token.string == "hh")
                        {
                            part.type = template_part::template_type::hour_leading_zero;
                            break;
                        }

                        throw std::runtime_error("unhandled");
                        break;

                    case 's':
                        if (token.string == "s")
                        {
                            part.type = template_part::template_type::second;
                            break;
                        }
                        else if (token.string == "ss")
                        {
                            part.type = template_part::template_type::second_leading_zero;
                            break;
                        }

                        throw std::runtime_error("unhandled");
                        break;

                    case 'A':
                        section.twelve_hour = true;

                        if (token.string == "AM/PM")
                        {
                            part.type = template_part::template_type::am_pm;
                            break;
                        }
                        else if (token.string == "A/P")
                        {
                            part.type = template_part::template_type::a_p;
                            break;
                        }

                        throw std::runtime_error("unhandled");
                        break;

                    default:
                        throw std::runtime_error("unhandled");
                        break;
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
                std::size_t exponent_index = 0;

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
                        exponent_index = i;
                    }

                    if (part.placeholders.percentage)
                    {
                        percentage = true;
                    }

                    if (part.type == template_part::template_type::second || part.type == template_part::template_type::second_leading_zero)
                    {
                        seconds = true;
                        seconds_index = i;
                    }

                    if (seconds && part.placeholders.type == format_placeholders::placeholders_type::fractional_part)
                    {
                        fractional_seconds = true;
                    }

                    // TODO this block needs improvement
                    if (part.type == template_part::template_type::month_number || part.type == template_part::template_type::month_number_leading_zero)
                    {
                        if (code.parts.size() > 1 && i < code.parts.size() - 2)
                        {
                            const auto &next = code.parts[i + 1];
                            const auto &after_next = code.parts[i + 2];

                            if ((next.type == template_part::template_type::second || next.type == template_part::template_type::second_leading_zero) || (next.type == template_part::template_type::text && next.string == ":" && (after_next.type == template_part::template_type::second || after_next.type == template_part::template_type::second_leading_zero)))
                            {
                                fix = true;
                                leading_zero = part.type == template_part::template_type::month_number_leading_zero;
                                minutes_index = i;
                            }
                        }

                        if (!fix && i > 1)
                        {
                            const auto &previous = code.parts[i - 1];
                            const auto &before_previous = code.parts[i - 2];

                            if (previous.type == template_part::template_type::text && previous.string == ":" && (before_previous.type == template_part::template_type::hour_leading_zero || before_previous.type == template_part::template_type::hour))
                            {
                                fix = true;
                                leading_zero = part.type == template_part::template_type::month_number_leading_zero;
                                minutes_index = i;
                            }
                        }
                    }
                }

                if (fix)
                {
                    code.parts[minutes_index].type = leading_zero ? template_part::template_type::minute_leading_zero
                                                                  : template_part::template_type::minute;
                }

                if (integer_part && !fractional_part)
                {
                    code.parts[integer_part_index].placeholders.type = format_placeholders::placeholders_type::integer_only;
                }

                if (integer_part && fractional_part && percentage)
                {
                    code.parts[integer_part_index].placeholders.percentage = true;
                }

                if (exponent)
                {
                    const auto &next = code.parts[exponent_index + 1];
                    auto temp = code.parts[exponent_index].placeholders.type;
                    code.parts[exponent_index].placeholders = next.placeholders;
                    code.parts[exponent_index].placeholders.type = temp;
                    code.parts.erase(code.parts.begin() + static_cast<std::ptrdiff_t>(exponent_index + 1));

                    for (std::size_t i = 0; i < code.parts.size(); ++i)
                    {
                        code.parts[i].placeholders.scientific = true;
                    }
                }

                if (fraction)
                {
                    code.parts[fraction_numerator_index].placeholders.type = format_placeholders::placeholders_type::fraction_numerator;
                    code.parts[fraction_denominator_index].placeholders.type = format_placeholders::placeholders_type::fraction_denominator;

                    for (std::size_t i = 0; i < code.parts.size(); ++i)
                    {
                        if (code.parts[i].placeholders.type == format_placeholders::placeholders_type::integer_part)
                        {
                            code.parts[i].placeholders.type = format_placeholders::placeholders_type::fraction_integer;
                        }
                    }
                }

                if (fractional_seconds)
                {
                    if (code.parts[seconds_index].type == template_part::template_type::second)
                    {
                        code.parts[seconds_index].type = template_part::template_type::second_fractional;
                    }
                    else
                    {
                        code.parts[seconds_index].type = template_part::template_type::second_leading_zero_fractional;
                    }
                }
            }

            if (_format.size() > 4)
            {
                throw std::runtime_error("too many format codes");
            }

            if (_format.size() > 2)
            {
                if (_format[0].has_condition && _format[1].has_condition && _format[2].has_condition)
                {
                    throw std::runtime_error(
                        "format should have a maximum of two codes with conditions");
                }
            }
        }

        std::vector<number_format_token> parse_tokens() const
        {
            auto to_lower = [](char c)
            {
                return static_cast<char>(std::tolower(static_cast<std::uint8_t>(c)));
            };

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
                    {
                        throw std::runtime_error("missing ]");
                    }

                    if (_format_string[position] == ']')
                    {
                        throw std::runtime_error("empty []");
                    }

                    do
                    {
                        token.string.push_back(_format_string[position++]);
                    } while (position < _format_string.size() && _format_string[position] != ']');

                    if (token.string[0] == '<' || token.string[0] == '>' || token.string[0] == '=')
                    {
                        token.type = number_format_token::token_type::condition;
                    }
                    else if (token.string[0] == '$')
                    {
                        token.type = number_format_token::token_type::locale;
                    }
                    else if (token.string.size() <= 2 && ((token.string == "h" || token.string == "hh") || (token.string == "m" || token.string == "mm") || (token.string == "s" || token.string == "ss")))
                    {
                        token.type = number_format_token::token_type::datetime;
                        token.string = "[" + token.string + "]";
                    }
                    else
                    {
                        token.type = number_format_token::token_type::color;
                    }

                    ++position;

                    break;

                case '\\':
                    token.type = number_format_token::token_type::text;
                    token.string.push_back(_format_string[position++]);

                    break;

                case 'G':
                    if (_format_string.substr(position - 1, 7) != "General")
                    {
                        throw std::runtime_error("expected General");
                    }

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
                    token.string.push_back(to_lower(current_char));

                    while (_format_string[position] == current_char)
                    {
                        token.string.push_back(to_lower(current_char));
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
                    {
                        throw std::runtime_error("expected AM/PM or A/P");
                    }

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
                    {
                        token.string.append(_format_string.substr(start, end - start));
                    }

                    position = end + 1;

                    break;
                }

                case ';':
                    token.type = number_format_token::token_type::end_section;
                    break;

                case '(':
                    token.type = number_format_token::token_type::text;
                    token.string.push_back(current_char);

                    break;

                case ')':
                    token.type = number_format_token::token_type::text;
                    token.string.push_back(current_char);

                    break;

                case '-':
                    token.type = number_format_token::token_type::text;
                    token.string.push_back(current_char);

                    break;

                case '+':
                    token.type = number_format_token::token_type::text;
                    token.string.push_back(current_char);

                    break;

                case ':':
                    token.type = number_format_token::token_type::text;
                    token.string.push_back(current_char);

                    break;

                case ' ':
                    token.type = number_format_token::token_type::text;
                    token.string.push_back(current_char);

                    break;

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
                        break;
                    }

                    break;

                default:
                    throw std::runtime_error("unexpected character");
                }
                tokens.push_back(token);
            }
            if (tokens.size() > 0 && tokens[tokens.size() - 1].type != number_format_token::token_type::end_section)
            {
                number_format_token token;
                token.type = number_format_token::token_type::end_section;
                tokens.push_back(token);
            }

            return tokens;
        }
        format_placeholders parse_placeholders(const std::string &placeholders_string) const
        {
            format_placeholders p;

            if (placeholders_string == "General")
            {
                p.type = format_placeholders::placeholders_type::general;
                return p;
            }
            else if (placeholders_string == "@")
            {
                p.type = format_placeholders::placeholders_type::text;
                return p;
            }
            else if (placeholders_string.front() == '.')
            {
                p.type = format_placeholders::placeholders_type::fractional_part;
            }
            else if (placeholders_string.front() == 'E')
            {
                p.type = placeholders_string[1] == '+' ? format_placeholders::placeholders_type::scientific_exponent_plus : format_placeholders::placeholders_type::scientific_exponent_minus;
                return p;
            }
            else
            {
                p.type = format_placeholders::placeholders_type::integer_part;
            }

            if (placeholders_string.back() == '%')
            {
                p.percentage = true;
            }

            std::vector<std::size_t> comma_indices;

            for (std::size_t i = 0; i < placeholders_string.size(); ++i)
            {
                auto c = placeholders_string[i];

                if (c == '0')
                {
                    ++p.num_zeros;
                }
                else if (c == '#')
                {
                    ++p.num_optionals;
                }
                else if (c == '?')
                {
                    ++p.num_spaces;
                }
                else if (c == ',')
                {
                    comma_indices.push_back(i);
                }
            }

            if (!comma_indices.empty())
            {
                std::size_t i = placeholders_string.size() - 1;

                while (!comma_indices.empty() && i == comma_indices.back())
                {
                    ++p.thousands_scale;
                    --i;
                    comma_indices.pop_back();
                }

                p.use_comma_separator = !comma_indices.empty();
            }

            return p;
        }

        std::string format(double number, bool is_date1904 = false)
        {
            if (_format[0].has_condition)
            {
                if (_format[0].condition.satisfied_by(number))
                    return format_number(_format[0], number, is_date1904);
                if (_format.size() == 1)
                    return std::string(11, '#');
                if (!_format[1].has_condition || _format[1].condition.satisfied_by(number))
                    return format_number(_format[1], number, is_date1904);
                if (_format.size() == 2)
                    return std::string(11, '#');
                return format_number(_format[2], number, is_date1904);
            }

            // no conditions, format based on sign:

            // 1 section, use for all
            if (_format.size() == 1)
                return format_number(_format[0], number, is_date1904);
            // 2 sections, first for positive and zero, second for negative
            else if (_format.size() == 2)
            {
                if (number >= 0)
                    return format_number(_format[0], number, is_date1904);
                return format_number(_format[1], std::fabs(number), is_date1904);
            }
            // 3+ sections, first for positive, second for negative, third for zero
            else
            {
                if (number > 0)
                    return format_number(_format[0], number, is_date1904);
                if (number < 0)
                    return format_number(_format[1], std::fabs(number), is_date1904);
                return format_number(_format[2], number, is_date1904);
            }
        }
        std::string format(const std::string &text)
        {
            if (_format.size() < 4)
            {
                format_code temp;
                template_part temp_part;
                temp_part.type = template_part::template_type::general;
                temp_part.placeholders.type = format_placeholders::placeholders_type::general;
                temp.parts.push_back(temp_part);
                return format_text(temp, text);
            }

            return format_text(_format[3], text);
        }

    private:
        std::string fill_placeholders(const format_placeholders &p, double number)
        {
            std::string result;

            if (p.type == format_placeholders::placeholders_type::general || p.type == format_placeholders::placeholders_type::text)
            {
                result = std::to_string(number);

                while (result.back() == '0')
                {
                    result.pop_back();
                }

                if (result.back() == '.')
                {
                    result.pop_back();
                }

                return result;
            }

            if (p.percentage)
            {
                number *= 100;
            }

            if (p.thousands_scale > 0)
            {
                number /= std::pow(1000.0, p.thousands_scale);
            }

            auto integer_part = static_cast<int>(number);

            if (p.type == format_placeholders::placeholders_type::integer_only || p.type == format_placeholders::placeholders_type::integer_part || p.type == format_placeholders::placeholders_type::fraction_integer)
            {
                result = std::to_string(integer_part);

                while (result.size() < p.num_zeros)
                {
                    result = "0" + result;
                }

                while (result.size() < p.num_zeros + p.num_spaces)
                {
                    result = " " + result;
                }

                if (p.use_comma_separator)
                {
                    std::vector<char> digits(result.rbegin(), result.rend());
                    std::string temp;

                    for (std::size_t i = 0; i < digits.size(); i++)
                    {
                        temp.push_back(digits[i]);

                        if (i % 3 == 2)
                        {
                            temp.push_back(',');
                        }
                    }

                    result = std::string(temp.rbegin(), temp.rend());
                }

                if (p.percentage && p.type == format_placeholders::placeholders_type::integer_only)
                {
                    result.push_back('%');
                }
            }
            else if (p.type == format_placeholders::placeholders_type::fractional_part)
            {
                auto fractional_part = number - integer_part;
                result = std::fabs(fractional_part) < std::numeric_limits<double>::min()
                             ? std::string(".")
                             : std::to_string(fractional_part).substr(1);

                while (result.back() == '0' || result.size() > (p.num_zeros + p.num_optionals + p.num_spaces + 1))
                {
                    result.pop_back();
                }

                while (result.size() < p.num_zeros + 1)
                {
                    result.push_back('0');
                }

                while (result.size() < p.num_zeros + p.num_optionals + p.num_spaces + 1)
                {
                    result.push_back(' ');
                }

                if (p.percentage)
                {
                    result.push_back('%');
                }
            }

            return result;
        }
        std::string fill_fraction_placeholders(const format_placeholders &numerator, const format_placeholders &denominator, double number, bool improper)
        {
            auto fractional_part = number - static_cast<int>(number);
            auto original_fractional_part = fractional_part;
            fractional_part *= 10;

            while (std::abs(fractional_part - static_cast<int>(fractional_part)) > 0.000001 && std::abs(fractional_part - static_cast<int>(fractional_part)) < 0.999999)
            {
                fractional_part *= 10;
            }

            fractional_part = static_cast<int>(fractional_part);
            auto denominator_digits = denominator.num_zeros + denominator.num_optionals + denominator.num_spaces;
            //    auto denominator_digits =
            //    static_cast<std::size_t>(std::ceil(std::log10(fractional_part)));

            auto lower = static_cast<int>(std::pow(10, denominator_digits - 1));
            auto upper = static_cast<int>(std::pow(10, denominator_digits));
            auto best_denominator = lower;
            auto best_difference = 1000.0;

            for (int i = lower; i < upper; ++i)
            {
                auto numerator_full = original_fractional_part * i;
                auto numerator_rounded = static_cast<int>(std::round(numerator_full));
                auto difference = std::fabs(original_fractional_part - (numerator_rounded / static_cast<double>(i)));

                if (difference < best_difference)
                {
                    best_difference = difference;
                    best_denominator = i;
                }
            }

            auto numerator_rounded = static_cast<int>(std::round(original_fractional_part * best_denominator));
            return std::to_string(numerator_rounded) + "/" + std::to_string(best_denominator);
        }
        std::string fill_scientific_placeholders(const format_placeholders &integer_part, const format_placeholders &fractional_part, const format_placeholders &exponent_part, double number)
        {
            std::size_t logarithm = 0;

            if (number != 0.0)
            {
                logarithm = static_cast<std::size_t>(std::log10(number));

                if (integer_part.num_zeros + integer_part.num_optionals > 1)
                {
                    logarithm = integer_part.num_zeros + integer_part.num_optionals;
                }
            }

            number /= std::pow(10.0, logarithm);

            auto integer = static_cast<int>(number);
            auto fraction = number - integer;

            std::string integer_string = std::to_string(integer);

            if (number == 0.0)
            {
                integer_string = std::string(integer_part.num_zeros + integer_part.num_optionals, '0');
            }

            std::string fractional_string = std::to_string(fraction).substr(1);

            while (fractional_string.size() > fractional_part.num_zeros + fractional_part.num_optionals + 1)
            {
                fractional_string.pop_back();
            }

            std::string exponent_string = std::to_string(logarithm);

            while (exponent_string.size() < fractional_part.num_zeros)
            {
                exponent_string.insert(0, "0");
            }

            if (exponent_part.type == format_placeholders::placeholders_type::scientific_exponent_plus)
            {
                exponent_string.insert(0, "E+");
            }
            else
            {
                exponent_string.insert(0, "E");
            }

            return integer_string + fractional_string + exponent_string;
        }
        std::string format_number(const format_code &format, double number, bool is_date1904 = false)
        {
            static const std::vector<std::string> month_names = std::vector<std::string>{"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
            static const std::vector<std::string> day_names = std::vector<std::string>{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};

            std::string result;

            if (number < 0)
            {
                result.push_back('-');

                if (format.is_datetime)
                {
                    return std::string(11, '#');
                }
            }

            number = std::fabs(number);

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
                {
                    result.push_back(' ');
                    break;
                }

                case template_part::template_type::text:
                {
                    result.append(part.string);
                    break;
                }

                case template_part::template_type::fill:
                {
                    fill = true;
                    fill_index = result.size();
                    fill_character = part.string;
                    break;
                }

                case template_part::template_type::general:
                {
                    if (part.placeholders.type == format_placeholders::placeholders_type::fractional_part && (format.is_datetime || format.is_timedelta))
                    {
                        auto digits = std::min(static_cast<std::size_t>(6), part.placeholders.num_zeros + part.placeholders.num_optionals);
                        auto denominator = static_cast<int>(std::pow(10.0, digits));
                        auto fractional_seconds = dt.microsecond / 1.0E6 * denominator;
                        fractional_seconds = std::round(fractional_seconds) / denominator;
                        result.append(
                            fill_placeholders(part.placeholders, fractional_seconds));
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
                {
                    result.append(std::to_string(dt.day));
                    break;
                }

                case template_part::template_type::day_number_leading_zero:
                {
                    if (dt.day < 10)
                    {
                        result.push_back('0');
                    }

                    result.append(std::to_string(dt.day));
                    break;
                }

                case template_part::template_type::month_abbreviation:
                {
                    result.append(month_names.at(static_cast<std::size_t>(dt.month) - 1)
                                      .substr(0, 3));
                    break;
                }

                case template_part::template_type::month_name:
                {
                    result.append(month_names.at(static_cast<std::size_t>(dt.month) - 1));
                    break;
                }

                case template_part::template_type::month_number:
                {
                    result.append(std::to_string(dt.month));
                    break;
                }

                case template_part::template_type::month_number_leading_zero:
                {
                    if (dt.month < 10)
                    {
                        result.push_back('0');
                    }

                    result.append(std::to_string(dt.month));
                    break;
                }

                case template_part::template_type::year_short:
                {
                    if (dt.year % 1000 < 10)
                    {
                        result.push_back('0');
                    }

                    result.append(std::to_string(dt.year % 1000));
                    break;
                }

                case template_part::template_type::year_long:
                {
                    result.append(std::to_string(dt.year));
                    break;
                }

                case template_part::template_type::hour:
                {
                    result.append(std::to_string(hour));
                    break;
                }

                case template_part::template_type::hour_leading_zero:
                {
                    if (hour < 10)
                    {
                        result.push_back('0');
                    }

                    result.append(std::to_string(hour));
                    break;
                }

                case template_part::template_type::minute:
                {
                    result.append(std::to_string(dt.minute));
                    break;
                }

                case template_part::template_type::minute_leading_zero:
                {
                    if (dt.minute < 10)
                    {
                        result.push_back('0');
                    }

                    result.append(std::to_string(dt.minute));
                    break;
                }

                case template_part::template_type::second:
                {
                    result.append(
                        std::to_string(dt.second + (dt.microsecond > 500000 ? 1 : 0)));
                    break;
                }

                case template_part::template_type::second_fractional:
                {
                    result.append(std::to_string(dt.second));
                    break;
                }

                case template_part::template_type::second_leading_zero:
                {
                    if ((dt.second + (dt.microsecond > 500000 ? 1 : 0)) < 10)
                    {
                        result.push_back('0');
                    }

                    result.append(
                        std::to_string(dt.second + (dt.microsecond > 500000 ? 1 : 0)));
                    break;
                }

                case template_part::template_type::second_leading_zero_fractional:
                {
                    if (dt.second < 10)
                    {
                        result.push_back('0');
                    }

                    result.append(std::to_string(dt.second));
                    break;
                }

                case template_part::template_type::am_pm:
                {
                    if (dt.hour < 12)
                    {
                        result.append("AM");
                    }
                    else
                    {
                        result.append("PM");
                    }

                    break;
                }

                case template_part::template_type::a_p:
                {
                    if (dt.hour < 12)
                    {
                        result.append("A");
                    }
                    else
                    {
                        result.append("P");
                    }

                    break;
                }

                case template_part::template_type::elapsed_hours:
                {
                    result.append(std::to_string(24 * static_cast<int>(number) + dt.hour));
                    break;
                }

                case template_part::template_type::elapsed_minutes:
                {
                    result.append(std::to_string(24 * 60 * static_cast<int>(number) + (60 * dt.hour) + dt.minute));
                    break;
                }

                case template_part::template_type::elapsed_seconds:
                {
                    result.append(std::to_string(24 * 60 * 60 * static_cast<int>(number) + (60 * 60 * dt.hour) + (60 * dt.minute) + dt.second));
                    break;
                }

                case template_part::template_type::month_letter:
                {
                    result.append(month_names.at(static_cast<std::size_t>(dt.month) - 1)
                                      .substr(0, 1));
                    break;
                }

                case template_part::template_type::day_abbreviation:
                {
                    result.append(
                        day_names.at(static_cast<std::size_t>(dt.weekday())).substr(0, 3));
                    break;
                }

                case template_part::template_type::day_name:
                {
                    result.append(day_names.at(static_cast<std::size_t>(dt.weekday())));
                    break;
                }
                }
            }

            const std::size_t width = 11;

            if (fill && result.size() < width)
            {
                auto remaining = width - result.size();

                std::string fill_string(remaining, fill_character.front());
                // TODO: A UTF-8 character could be multiple bytes

                result = result.substr(0, fill_index) + fill_string + result.substr(fill_index);
            }

            return result;
        }
        std::string format_text(const format_code &format, const std::string &text)
        {
            std::string result;
            bool any_text_part = false;

            for (const auto &part : format.parts)
            {
                if (part.type == template_part::template_type::text)
                {
                    result.append(part.string);
                    any_text_part = true;
                }
                else if (part.type == template_part::template_type::general)
                {
                    if (part.placeholders.type == format_placeholders::placeholders_type::general || part.placeholders.type == format_placeholders::placeholders_type::text)
                    {
                        result.append(text);
                        any_text_part = true;
                    }
                }
            }

            if (!format.parts.empty() && !any_text_part)
            {
                return text;
            }

            return result;
        }
    };

    class reference
    {
    public:
        unsigned row;
        unsigned col;

        reference() noexcept : reference(0, 0) {}
        reference(unsigned row, unsigned col) noexcept : row(row), col(col) {}
        reference(const std::string &value) noexcept { this->value(value); }

        void value(const std::string &value) noexcept
        {
            row = col = 0;
            for (std::string::size_type i = 0; i < value.size(); ++i)
            {
                auto c = value[i];
                if (row == 0 && 'A' <= c && c <= 'Z')
                    col = col * 26 + (c - 'A') + 1;
                else if (col > 0 && '0' <= c && c <= '9')
                    row = row * 10 + (c - '0');
                else
                {
                    row = col = 0;
                    break;
                }
            }
        }
        std::string value() const noexcept
        {
            std::string value = "";
            if (row > 0 && col > 0)
            {
                auto col_ = col;
                while (col_ > 0)
                {
                    char c = (col_ - 1) % 26 + 'A';
                    value = c + value;
                    col_ = (col_ - (c - 'A' + 1)) / 26;
                }
                value += std::to_string(row);
            }
            return value;
        }

        operator bool() const noexcept { return row > 0 && col > 0; }
    };

    class cell
    {
    public:
        reference refer;
        std::string value;
        cell(reference reference, std::string value = "") noexcept : refer(reference), value(value) {}
        cell(std::string reference, std::string value = "") noexcept : refer(reference), value(value) {}
        cell(unsigned row, unsigned col, std::string value = "") noexcept : refer(row, col), value(value) {}
    };

    class worksheet
    {
    protected:
        struct package
        {
            mz_zip_archive archive{};

            bool date1904 = false;
            std::vector<std::string> shared_strings{}; // text
            std::map<unsigned, std::string> numfmts =  // id code
                {{0, "General"},
                 {1, "0"},
                 {2, "0.00"},
                 {3, "#,##0"},
                 {4, "#,##0.00"},
                 {9, "0%"},
                 {10, "0.00%"},
                 {11, "0.00E+00"},
                 {12, "# ?/?"},
                 {13, "# ?\?/??"},
                 {14, "mm-dd-yy"},
                 {15, "d-mmm-yy"},
                 {16, "d-mmm"},
                 {17, "mmm-yy"},
                 {18, "h:mm AM/PM"},
                 {19, "h:mm:ss AM/PM"},
                 {20, "h:mm"},
                 {21, "h:mm:ss"},
                 {22, "m/d/yy h:mm"},
                 {37, "#,##0 ;(#,##0)"},
                 {38, "#,##0 ;[Red](#,##0)"},
                 {39, "#,##0.00;(#,##0.00)"},
                 {40, "#,##0.00;[Red](#,##0.00)"},
                 {45, "mm:ss"},
                 {46, "[h]:mm:ss"},
                 {47, "mmss.0"},
                 {48, "##0.0E+0"},
                 {49, "@"}};
            std::vector<unsigned> cell_xfs{}; // id
            std::map<std::string, number_format *> number_formats{};

            virtual ~package()
            {
                mz_zip_reader_end(&archive);
                for (auto kv : number_formats)
                    delete kv.second;
            }
            std::string format(std::string format_string, const std::string &text)
            {
                if (number_formats.find(format_string) == number_formats.end())
                    number_formats[format_string] = new number_format(format_string);
                return number_formats[format_string]->format(text);
            }
            std::string format(std::string format_string, double number)
            {
                if (number_formats.find(format_string) == number_formats.end())
                    number_formats[format_string] = new number_format(format_string);
                return number_formats[format_string]->format(number, date1904);
            }
        };

    protected:
        std::shared_ptr<package> _package;

    private:
        std::string _part;
        std::string _name;
        std::vector<std::tuple<reference, reference, std::string>> _merge_cells;
        std::vector<std::vector<cell>> _rows;

    protected:
        worksheet(std::shared_ptr<package> package) noexcept : _package(package) {}
        worksheet(const std::string &name, const std::string &part, std::shared_ptr<package> package) noexcept : _name(name), _part(part), _package(package) {}
        static worksheet create(const std::string &name, const std::string &part, std::shared_ptr<package> package) noexcept { return worksheet(name, part, package); }

    private:
        std::string read_value(const std::string &v, const std::string &t, const std::string &s, std::string &error) const
        {
            if (t == "n" || t == "str" || t == "inlineStr")
            {
                return v;
            }
            else if (t == "b")
            {
                return v == "0" ? "FALSE" : "TRUE";
            }
            else if (t == "s")
            {
                auto index = std::stol(v);
                return index < _package->shared_strings.size() ? _package->shared_strings[index] : "";
            }
            else if (t == "d")
            {
                error = "date type is not supported";
            }
            else if (t == "e")
            {
                error = "cell error";
            }
            else
            {
                if (s == "")
                    return v;

                auto index = std::stol(s);
                auto xf = index < _package->cell_xfs.size() ? _package->cell_xfs[index] : 0;
                const auto &format = _package->numfmts.find(xf) != _package->numfmts.end() ? _package->numfmts[xf] : "";
                return _package->format(format, v);
            }
            return "";
        }

    public:
        std::string name() const noexcept { return _name; }
        std::map<std::string, std::string> read()
        {
            _merge_cells.clear();
            _rows.clear();

            std::map<std::string, std::string> errors;

            void *buffer = nullptr;
            size_t size = 0;
            if ((buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, _part.c_str(), &size, 0)) != nullptr)
            {
                /**
                 * <xsd:simpleType name="ST_Xstring">
                 *     <xsd:restriction base="xsd:string"/>
                 * </xsd:simpleType>
                 * <xsd:complexType name="CT_RElt">
                 *     <xsd:sequence>
                 *         <xsd:element name="t" type="s:ST_Xstring" minOccurs="1" maxOccurs="1"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Rst">
                 *     <xsd:sequence>
                 *         <xsd:element name="t" type="s:ST_Xstring" minOccurs="0" maxOccurs="1"/>
                 *         <xsd:element name="r" type="CT_RElt" minOccurs="0" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:simpleType name="ST_CellRef">
                 *     <xsd:restriction base="xsd:string"/>
                 * </xsd:simpleType>
                 * <xsd:simpleType name="ST_CellType">
                 *     <xsd:restriction base="xsd:string">
                 *         <xsd:enumeration value="b"/>
                 *         <xsd:enumeration value="d"/>
                 *         <xsd:enumeration value="n"/>
                 *         <xsd:enumeration value="e"/>
                 *         <xsd:enumeration value="s"/>
                 *         <xsd:enumeration value="str"/>
                 *         <xsd:enumeration value="inlineStr"/>
                 *     </xsd:restriction>
                 * </xsd:simpleType>
                 * <xsd:complexType name="CT_Cell">
                 *     <xsd:sequence>
                 *         <xsd:element name="f" type="CT_CellFormula" minOccurs="0" maxOccurs="1"/>
                 *         <xsd:element name="v" type="s:ST_Xstring" minOccurs="0" maxOccurs="1"/>
                 *         <xsd:element name="is" type="CT_Rst" minOccurs="0" maxOccurs="1"/>
                 *     </xsd:sequence>
                 *     <xsd:attribute name="r" type="ST_CellRef" use="optional"/>
                 *     <xsd:attribute name="s" type="xsd:unsignedInt" use="optional" default="0"/>
                 *     <xsd:attribute name="t" type="ST_CellType" use="optional" default="n"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Row">
                 *     <xsd:sequence>
                 *         <xsd:element name="c" type="CT_Cell" minOccurs="0" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 *     <xsd:attribute name="r" type="xsd:unsignedInt" use="optional"/>
                 *     <xsd:attribute name="spans" type="ST_CellSpans" use="optional"/>
                 *     <xsd:attribute name="s" type="xsd:unsignedInt" use="optional" default="0"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_SheetData">
                 *     <xsd:sequence>
                 *         <xsd:element name="row" type="CT_Row" minOccurs="0" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:simpleType name="ST_Ref">
                 *     <xsd:restriction base="xsd:string"/>
                 * </xsd:simpleType>
                 * <xsd:complexType name="CT_MergeCell">
                 *     <xsd:attribute name="ref" type="ST_Ref" use="required"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_MergeCells">
                 *     <xsd:sequence>
                 *         <xsd:element name="mergeCell" type="CT_MergeCell" minOccurs="1" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 *     <xsd:attribute name="count" type="xsd:unsignedInt" use="optional"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Worksheet">
                 *     <xsd:sequence>
                 *         <xsd:element name="sheetData" type="CT_SheetData" minOccurs="1" maxOccurs="1"/>
                 *         <xsd:element name="mergeCells" type="CT_MergeCells" minOccurs="0" maxOccurs="1"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:element name="worksheet" type="CT_Worksheet"/>
                 *
                 * <worksheet>
                 *     <sheetData>
                 *         <row r="1">
                 *             <c r="A1" s="11"><v>2</v></c>
                 *             <c r="B1" s="11"><v>3</v></c>
                 *             <c r="C1" s="11"><v>4</v></c>
                 *             <c r="D1" t="s"><v>0</v></c>
                 *             <c r="E1" t="inlineStr"><is><t>This is inline string example</t></is></c>
                 *             <c r="D1" t="d"><v>1976-11-22T08:30</v></c>
                 *             <c r="G1"><f>SUM(A1:A3)</f><v>9</v></c>
                 *             <c r="H1" s="11"/>
                 *         </row>
                 *     </sheetData>
                 *     <mergeCells count="5">
                 *         <mergeCell ref="A1:B2"/>
                 *         <mergeCell ref="C1:E5"/>
                 *         <mergeCell ref="A3:B6"/>
                 *         <mergeCell ref="A7:C7"/>
                 *         <mergeCell ref="A8:XFD9"/>
                 *     </mergeCells>
                 * <worksheet>
                 */
                pugi::xml_document doc;
                auto result = doc.load_buffer_inplace_own(buffer, size);
                if (!result)
                {
                    errors[_name] = "workseet open failed";
                    return errors;
                }

                for (auto mc = doc.child("worksheet").child("mergeCells").child("mergeCell"); mc; mc = mc.next_sibling("mergeCell"))
                {
                    auto ref = mc.attribute("ref");
                    if (ref)
                    {
                        std::string refs = ref.value();
                        auto split = refs.find(':');
                        if (split != std::string::npos && split < refs.size() - 1)
                            _merge_cells.push_back({refs.substr(0, split), refs.substr(split + 1), ""});
                    }
                }
                unsigned row_index = 0, col_index = 0;
                for (auto row = doc.child("worksheet").child("sheetData").child("row"); row; row = row.next_sibling("row"))
                {
                    std::string r = row.attribute("r").value();
                    row_index = r == "" ? row_index + 1 : std::stol(r);
                    col_index = 0;

                    std::vector<cell> cells;
                    for (auto c = row.child("c"); c; c = c.next_sibling("c"))
                    {
                        reference refer(c.attribute("r").value()); // "r" is optional
                        if (!refer)
                        {
                            refer.row = row_index;
                            refer.col = ++col_index;
                        }
                        col_index = refer.col;

                        if (refer.row != row_index) // Error in Microsoft Excel
                            continue;

                        std::string v = c.child("v").text().get(), t = c.attribute("t").value(), s = c.attribute("s").value();
                        /**
                         * (Ecma Office Open XML Part 1)
                         *
                         * when the cell's type t is inlineStr then only the element is is allowed as a child element.
                         *
                         * Cell containing an (inline) rich string, i.e., one not in the shared string table. If this cell type is used, then the cell value is in the is element rather than the v element in the cell (c element).
                         */
                        if (t == "inlineStr")
                        {
                            v = "";
                            auto is = c.child("is");
                            auto r = is.child("r");
                            if (r)
                            {
                                for (; r; r = r.next_sibling("r"))
                                    v += r.child("t").text().get();
                            }
                            else
                                v = is.child("t").text().get();
                        }

                        std::string error, value = c.child("f") ? v : read_value(v, t, s, error);
                        if (error != "")
                            errors[refer.value()] = error;
                        cells.push_back(cell(refer, value));
                    }

                    if (cells.size())
                        _rows.push_back(std::move(cells));
                }
            }
            return errors;
        }

        const std::vector<std::vector<cell>> &rows() const noexcept { return _rows; }
        std::vector<std::vector<cell>>::const_iterator begin() const noexcept { return _rows.begin(); }
        std::vector<std::vector<cell>>::const_iterator end() const noexcept { return _rows.end(); }
    };

    class workbook : private worksheet
    {
    private:
        const std::string _path;
        std::vector<worksheet> _worksheets;

    public:
        workbook(const std::string &path) noexcept : worksheet(std::make_shared<package>()), _path(path) {}

        bool read() noexcept
        {
            _worksheets.clear();

            if (!mz_zip_reader_init_file(&_package->archive, _path.c_str(), 0))
                return false;

            std::string workbook_part = "xl/workbook.xml";
            std::string shared_strings_part = "xl/sharedStrings.xml";
            std::string styles_part = "xl/styles.xml";

            std::map<std::string, std::string> sheets; // Id Target

            size_t size = 0;
            void *buffer = nullptr;

            pugi::xml_document doc;
            if ((buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, "_rels/.rels", &size, 0)) != nullptr)
            {
                /**
                 * <xsd:complexType name="CT_Relationship">
                 *     <xsd:simpleContent>
                 *         <xsd:extension base="xsd:string">
                 *             <xsd:attribute name="Target" type="xsd:anyURI" use="required"/>
                 *             <xsd:attribute name="Type" type="xsd:anyURI" use="required"/>
                 *             <xsd:attribute name="Id" type="xsd:ID" use="required"/>
                 *         </xsd:extension>
                 *     </xsd:simpleContent>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Relationships">
                 *     <xsd:sequence>
                 *         <xsd:element ref="Relationship" minOccurs="0" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:element name="Relationship" type="CT_Relationship"/>
                 * <xsd:element name="Relationships" type="CT_Relationships"/>
                 *
                 * <Relationships>
                 *     <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                 * </Relationships>
                 */
                auto result = doc.load_buffer_inplace_own(buffer, size);
                if (!result)
                    return false;

                for (auto res = doc.child("Relationships").child("Relationship"); res; res = res.next_sibling("Relationship"))
                {
                    auto type = res.attribute("Type"), target = res.attribute("Target");
                    if (type && target && !std::strcmp(type.value(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"))
                    {
                        workbook_part = target.value();
                        break;
                    }
                }
            }
            if ((buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, "xl/_rels/workbook.xml.rels", &size, 0)) != nullptr)
            {
                /**
                 * <xsd:complexType name="CT_Relationship">
                 *     <xsd:simpleContent>
                 *         <xsd:extension base="xsd:string">
                 *             <xsd:attribute name="Target" type="xsd:anyURI" use="required"/>
                 *             <xsd:attribute name="Type" type="xsd:anyURI" use="required"/>
                 *             <xsd:attribute name="Id" type="xsd:ID" use="required"/>
                 *         </xsd:extension>
                 *     </xsd:simpleContent>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Relationships">
                 *     <xsd:sequence>
                 *         <xsd:element ref="Relationship" minOccurs="0" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:element name="Relationship" type="CT_Relationship"/>
                 * <xsd:element name="Relationships" type="CT_Relationships"/>
                 *
                 * <Relationships>
                 *     <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                 *     <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                 *     <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
                 *     <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                 * </Relationships>
                 */
                auto result = doc.load_buffer_inplace_own(buffer, size);
                if (!result)
                    return false;

                shared_strings_part = styles_part = "";
                for (auto res = doc.child("Relationships").child("Relationship"); res; res = res.next_sibling("Relationship"))
                {
                    auto id = res.attribute("Id"), type = res.attribute("Type"), target = res.attribute("Target");
                    if (id && type && target)
                    {
                        if (!std::strcmp(type.value(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"))
                            shared_strings_part = std::string("xl/") + target.value();
                        else if (!std::strcmp(type.value(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"))
                            styles_part = std::string("xl/") + target.value();
                        else if (!std::strcmp(type.value(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"))
                            sheets[id.value()] = std::string("xl/") + target.value();
                    }
                }
            }
            if (shared_strings_part != "" && (buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, shared_strings_part.c_str(), &size, 0)) != nullptr)
            {
                /**
                 * <xsd:simpleType name="ST_Xstring">
                 *     <xsd:restriction base="xsd:string"/>
                 * </xsd:simpleType>
                 * <xsd:complexType name="CT_RElt">
                 *     <xsd:sequence>
                 *         <xsd:element name="t" type="s:ST_Xstring" minOccurs="1" maxOccurs="1"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Rst">
                 *     <xsd:sequence>
                 *         <xsd:element name="t" type="s:ST_Xstring" minOccurs="0" maxOccurs="1"/>
                 *         <xsd:element name="r" type="CT_RElt" minOccurs="0" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Sst">
                 *     <xsd:sequence>
                 *         <xsd:element name="si" type="CT_Rst" minOccurs="0" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 *     <xsd:attribute name="count" type="xsd:unsignedInt" use="optional"/>
                 *     <xsd:attribute name="uniqueCount" type="xsd:unsignedInt" use="optional"/>
                 * </xsd:complexType>
                 * <xsd:element name="sst" type="CT_Sst"/>
                 *
                 * <sst count="2" uniqueCount="2">
                 *     <si><t>23  &#10;        &#10;         as</t></si>
                 *     <si>
                 *         <r><t>a</t></r>
                 *         <r><t>b</t></r>
                 *         <r><t>c</t></r>
                 *     </si>
                 *     <si><t>cd</t></si>
                 * </sst>
                 */
                auto result = doc.load_buffer_inplace_own(buffer, size);
                if (!result)
                    return false;

                for (auto si = doc.child("sst").child("si"); si; si = si.next_sibling("si"))
                {
                    std::string t;
                    auto r = si.child("r");
                    if (r)
                    {
                        for (; r; r = r.next_sibling("r"))
                            t += r.child("t").text().get();
                    }
                    else
                        t = si.child("t").text().get();
                    _package->shared_strings.push_back(t);
                }
            }
            if (styles_part != "" && (buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, styles_part.c_str(), &size, 0)) != nullptr)
            {
                /**
                 * <xsd:simpleType name="ST_NumFmtId">
                 *     <xsd:restriction base="xsd:unsignedInt"/>
                 * </xsd:simpleType>
                 * <xsd:simpleType name="ST_Xstring">
                 *     <xsd:restriction base="xsd:string"/>
                 * </xsd:simpleType>
                 * <xsd:complexType name="CT_NumFmt">
                 *     <xsd:attribute name="numFmtId" type="ST_NumFmtId" use="required"/>
                 *     <xsd:attribute name="formatCode" type="s:ST_Xstring" use="required"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_NumFmts">
                 *     <xsd:sequence>
                 *         <xsd:element name="numFmt" type="CT_NumFmt" minOccurs="0" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 *     <xsd:attribute name="count" type="xsd:unsignedInt" use="optional"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Xf">
                 *     <xsd:attribute name="numFmtId" type="ST_NumFmtId" use="optional"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_CellXfs">
                 *     <xsd:sequence>
                 *         <xsd:element name="xf" type="CT_Xf" minOccurs="1" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 *     <xsd:attribute name="count" type="xsd:unsignedInt" use="optional"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Stylesheet">
                 *     <xsd:sequence>
                 *         <xsd:element name="numFmts" type="CT_NumFmts" minOccurs="0" maxOccurs="1"/>
                 *         <xsd:element name="cellXfs" type="CT_CellXfs" minOccurs="0" maxOccurs="1"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:element name="styleSheet" type="CT_Stylesheet"/>
                 *
                 * <styleSheet>
                 *     <numFmts count="4">
                 *         <numFmt numFmtId="44" formatCode="_ &quot;￥&quot;* #,##0.00_ ;_ &quot;￥&quot;* \-#,##0.00_ ;_ &quot;￥&quot;* &quot;-&quot;??_ ;_ @_ "/>
                 *         <numFmt numFmtId="41" formatCode="_ * #,##0_ ;_ * \-#,##0_ ;_ * &quot;-&quot;_ ;_ @_ "/>
                 *         <numFmt numFmtId="43" formatCode="_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * &quot;-&quot;??_ ;_ @_ "/>
                 *         <numFmt numFmtId="42" formatCode="_ &quot;￥&quot;* #,##0_ ;_ &quot;￥&quot;* \-#,##0_ ;_ &quot;￥&quot;* &quot;-&quot;_ ;_ @_ "/>
                 *     </numFmts>
                 *     <cellXfs count="2">
                 *         <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                 *         <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"/>
                 *     </cellXfs>
                 * </styleSheet>
                 */
                auto result = doc.load_buffer_inplace_own(buffer, size);
                if (!result)
                    return false;

                auto style = doc.child("styleSheet");
                if (style)
                {
                    for (auto nf = style.child("numFmts").child("numFmt"); nf; nf = nf.next_sibling("numFmt"))
                    {
                        auto id = nf.attribute("numFmtId"), code = nf.attribute("formatCode");
                        if (id && code)
                            _package->numfmts[std::stol(id.value())] = code.value();
                    }
                    for (auto xf = style.child("cellXfs").child("xf"); xf; xf = xf.next_sibling("xf"))
                    {
                        auto id = xf.attribute("numFmtId");
                        if (id)
                            _package->cell_xfs.push_back(std::stol(id.value()));
                    }
                }
            }
            if ((buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, workbook_part.c_str(), &size, 0)) != nullptr)
            {
                /**
                 * <xsd:simpleType name="ST_Xstring">
                 *     <xsd:restriction base="xsd:string"/>
                 * </xsd:simpleType>
                 * <xsd:complexType name="CT_Sheet">
                 *     <xsd:attribute name="name" type="s:ST_Xstring" use="required"/>
                 *     <xsd:attribute name="sheetId" type="xsd:unsignedInt" use="required"/>
                 *     <xsd:attribute ref="r:id" use="required"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Sheets">
                 *     <xsd:sequence>
                 *         <xsd:element name="sheet" type="CT_Sheet" minOccurs="1" maxOccurs="unbounded"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_WorkbookPr">
                 *     <xsd:attribute name="date1904" type="xsd:boolean" use="optional" default="false"/>
                 * </xsd:complexType>
                 * <xsd:complexType name="CT_Workbook">
                 *     <xsd:sequence>
                 *         <xsd:element name="sheets" type="CT_Sheets" minOccurs="1" maxOccurs="1"/>
                 *         <xsd:element name="workbookPr" type="CT_WorkbookPr" minOccurs="0" maxOccurs="1"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:element name="workbook" type="CT_Workbook"/>
                 *
                 * <workbook>
                 *     <sheets>
                 *         <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                 *         <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
                 *     </sheets>
                 *     <workbookPr date1904="1"/>
                 * </workbook>
                 */
                auto result = doc.load_buffer_inplace_own(buffer, size);
                if (!result)
                    return false;
                auto workbook = doc.child("workbook");
                for (auto sheet = workbook.child("sheets").child("sheet"); sheet; sheet = sheet.next_sibling("sheet"))
                {
                    auto name = sheet.attribute("name"), rid = sheet.attribute("r:id");
                    if (name && rid && sheets.find(rid.value()) != sheets.end())
                    {
                        const auto &part = sheets[rid.value()];
                        if (mz_zip_reader_locate_file(&_package->archive, part.c_str(), nullptr, 0) != -1)
                            _worksheets.push_back(worksheet::create(name.value(), part, _package));
                    }
                }
                auto workbookPr = doc.child("workbookPr");
                if (workbookPr)
                {
                    auto date1904 = workbookPr.attribute("date1904");
                    auto date1904_value = date1904 ? date1904.value() : "0";
                    _package->date1904 = std::strcmp(date1904_value, "1") == 0 || std::strcmp(date1904_value, "true") == 0;
                }
            }

            return true;
        }

        const std::vector<worksheet> &worksheets() const noexcept { return _worksheets; }
        std::vector<worksheet>::const_iterator begin() const noexcept { return _worksheets.begin(); }
        std::vector<worksheet>::const_iterator end() const noexcept { return _worksheets.end(); }
    };
} // namespace xlsxtext