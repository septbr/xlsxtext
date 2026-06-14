#pragma once

#include <memory>
#include <string>

namespace xlsxtext
{

class number_format
{
public:
    number_format();
    explicit number_format(const std::string& format_string);
    ~number_format();

    number_format(number_format&&) noexcept;
    number_format& operator=(number_format&&) noexcept;

    number_format(const number_format&) = delete;
    number_format& operator=(const number_format&) = delete;

    std::string format(double number, bool date1904 = false) const;
    std::string format(const std::string& text) const;

private:
    struct impl;
    std::unique_ptr<impl> _impl;
};

} // namespace xlsxtext