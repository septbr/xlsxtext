#include <xlsxtext.hpp>

#include <iostream>
#include <cstring>
#include <ctime>
#include <cmath>
#include <iomanip>
#include <vector>
#include <functional>

#ifdef _WIN32
#include <windows.h>
#endif

struct TestResult
{
    int passed = 0;
    int failed = 0;
};

void test_number_format()
{
    TestResult total;
    
    auto test_case = [&](const std::string& format_str, double value, const std::string& expected) {
        try {
            xlsxtext::number_format fmt(format_str);
            std::string result = fmt.format(value);
            bool pass = (result == expected);
            if (pass) {
                total.passed++;
            } else {
                total.failed++;
                std::cout << "[FAIL] format: \"" << format_str << "\""
                          << ", value: " << value
                          << ", result: \"" << result << "\""
                          << ", expected: \"" << expected << "\""
                          << std::endl;
            }
        } catch (const std::exception& e) {
            total.failed++;
            std::cout << "[ERROR] format: \"" << format_str << "\" - " << e.what() << std::endl;
        }
    };

    auto test_text = [&](const std::string& format_str, const std::string& text, const std::string& expected) {
        try {
            xlsxtext::number_format fmt(format_str);
            std::string result = fmt.format(text);
            bool pass = (result == expected);
            if (pass) {
                total.passed++;
            } else {
                total.failed++;
                std::cout << "[FAIL] format: \"" << format_str << "\""
                          << ", text: \"" << text << "\""
                          << ", result: \"" << result << "\""
                          << ", expected: \"" << expected << "\""
                          << std::endl;
            }
        } catch (const std::exception& e) {
            total.failed++;
            std::cout << "[ERROR] format: \"" << format_str << "\" - " << e.what() << std::endl;
        }
    };

    std::cout << "=== number_format Comprehensive Tests ===" << std::endl << std::endl;

    // ========== General Format ==========
    test_case("General", 0, "0");
    test_case("General", 1, "1");
    test_case("General", -1, "-1");
    test_case("General", 12345.6789, "12345.6789");
    test_case("General", -12345.6789, "-12345.6789");
    test_case("General", 0.00001, "1e-05");
    test_case("General", 123456789012, "123456789012");
    test_case("General", 1234567890123, "1234567890123");  // May be scientific
    
    // ========== Integer Placeholders (0) ==========
    test_case("0", 0, "0");
    test_case("0", 5, "5");
    test_case("0", 123, "123");
    test_case("0", 123.456, "123");
    test_case("0", 123.5, "124");  // Rounding
    test_case("0", -123, "-123");
    test_case("00", 5, "05");
    test_case("00", 123, "123");
    test_case("00000", 123, "00123");
    test_case("00000", 0, "00000");
    test_case("00000", 99999, "99999");
    test_case("00000", 100000, "100000");
    
    // ========== Decimal Placeholders ==========
    test_case("0.0", 123.456, "123.5");
    test_case("0.00", 123.456, "123.46");
    test_case("0.000", 123.456, "123.456");
    test_case("0.0000", 123.456, "123.4560");
    test_case("0.00", 0.004, "0.00");  // Round down
    test_case("0.00", 0.005, "0.01");  // Round up
    test_case("0.00", 0.006, "0.01");
    test_case("0.00", 1.995, "2.00");  // Carry
    test_case("0.00", 9.999, "10.00"); // Carry overflow
    
    // ========== Optional Placeholders (#) ==========
    test_case("#", 0, "");
    test_case("#", 5, "5");
    test_case("#", 123, "123");
    test_case("#.##", 123.4, "123.4");
    test_case("#.##", 123.456, "123.46");
    test_case("#.##", 123, "123");
    test_case("#.##", 0.1, "0.1");
    test_case("##.##", 0.1, "0.1");
    test_case("#.00", 123.4, "123.40");
    test_case("0.##", 123.4, "123.4");
    test_case("0.##", 123, "123");
    
    // ========== Space Placeholder (?) ==========
    test_case("?.??", 1.2, "1.2 ");   // Space alignment
    test_case("?.??", 1.23, "1.23");
    test_case("?.??", 1.234, "1.23");
    test_case("0.??", 1.2, "1.2 ");
    test_case("??.??", 12.3, "12.3 ");
    
    // ========== Thousands Separator ==========
    test_case("#,##0", 0, "0");
    test_case("#,##0", 100, "100");
    test_case("#,##0", 1000, "1,000");
    test_case("#,##0", 10000, "10,000");
    test_case("#,##0", 100000, "100,000");
    test_case("#,##0", 1000000, "1,000,000");
    test_case("#,##0", 1234567, "1,234,567");
    test_case("#,##0", 1234567890, "1,234,567,890");
    test_case("#,##0.00", 1234.56, "1,234.56");
    test_case("#,##0.00", 1234567.89, "1,234,567.89");
    test_case("#,##0.00", -1234.56, "-1,234.56");
    
    // ========== Thousands Scaling ==========
    test_case("#,##0,", 1234, "1");
    test_case("#,##0,", 12345, "12");
    test_case("#,##0,", 123456, "123");
    test_case("#,##0,", 1234567, "1,235");  // Rounded
    test_case("#,##0,", 12345678, "12,346");
    test_case("#,##0,,", 123456789, "123");
    test_case("#,##0,,", 1234567890, "1,235");  // Rounded
    test_case("#,##0,,,", 1234567890123, "1,235");
    test_case("0,," , 123456789, "123");
    
    // ========== Percentage ==========
    test_case("0%", 0, "0%");
    test_case("0%", 0.5, "50%");
    test_case("0%", 1, "100%");
    test_case("0%", 0.01, "1%");
    test_case("0%", 1.5, "150%");
    test_case("0.00%", 0.5678, "56.78%");
    test_case("0.00%", 0.56789, "56.79%");  // Rounding
    test_case("0.000%", 0.12345, "12.345%");
    test_case("0%", 0.005, "1%");  // Rounding
    test_case("#.##%", 0.123, "12.3%");
    test_case("#,##0%", 1.2345, "123%");
    
    // ========== Scientific Notation ==========
    test_case("0.00E+00", 0, "0.00E+00");
    test_case("0.00E+00", 1, "1.00E+00");
    test_case("0.00E+00", 10, "1.00E+01");
    test_case("0.00E+00", 100, "1.00E+02");
    test_case("0.00E+00", 12345.6789, "1.23E+04");
    test_case("0.00E+00", 1234.56789, "1.23E+03");
    test_case("0.00E+00", 0.1, "1.00E-01");
    test_case("0.00E+00", 0.01, "1.00E-02");
    test_case("0.00E+00", 0.001, "1.00E-03");
    test_case("0.00E-00", 12345.6789, "1.23E04");
    test_case("0.00E-00", 0.001, "1.00E-03");
    test_case("#.##E+00", 1234, "1.23E+03");
    test_case("##0.0E+0", 1234, "1.2E+3");
    test_case("0.00E+00", -1234, "-1.23E+03");
    
    // ========== Date Formats ==========
    // Excel epoch: serial 1 = 1900-01-01
    test_case("yyyy-mm-dd", 1, "1900-01-01");
    test_case("yyyy-mm-dd", 2, "1900-01-02");
    test_case("yyyy-mm-dd", 59, "1900-02-28");
    test_case("yyyy-mm-dd", 60, "1900-02-29");  // Excel bug
    test_case("yyyy-mm-dd", 61, "1900-03-01");
    test_case("yyyy-mm-dd", 100, "1900-04-09");
    test_case("yyyy-mm-dd", 365, "1900-12-30");
    test_case("yyyy-mm-dd", 366, "1900-12-31");
    test_case("yyyy-mm-dd", 367, "1901-01-01");
    
    // 1970-01-01 (Unix epoch) = serial 25569
    test_case("yyyy-mm-dd", 25569, "1970-01-01");
    test_case("dddd", 25569, "Thursday");
    
    // 2024-01-01 = serial 45292
    test_case("yyyy-mm-dd", 45292, "2024-01-01");
    test_case("dddd", 45292, "Monday");
    test_case("dd/mm/yyyy", 45292, "01/01/2024");
    test_case("mm/dd/yyyy", 45292, "01/01/2024");
    test_case("dd-mm-yyyy", 45292, "01-01-2024");
    test_case("yyyy/mm/dd", 45292, "2024/01/01");
    
    // Various date components
    test_case("yy", 45292, "24");
    test_case("yyyy", 45292, "2024");
    test_case("m", 45292, "1");
    test_case("mm", 45292, "01");
    test_case("mmm", 45292, "Jan");
    test_case("mmmm", 45292, "January");
    test_case("d", 45292, "1");
    test_case("dd", 45292, "01");
    test_case("ddd", 45292, "Mon");
    test_case("dddd", 45292, "Monday");
    
    // Different months
    test_case("mmmm", 45292, "January");
    test_case("mmmm", 45323, "February");    // 2024-02-01
    test_case("mmmm", 45352, "March");       // 2024-03-01
    test_case("mmmm", 45383, "April");       // 2024-04-01
    test_case("mmmm", 45413, "May");         // 2024-05-01
    test_case("mmmm", 45444, "June");        // 2024-06-01
    test_case("mmmm", 45474, "July");        // 2024-07-01
    test_case("mmmm", 45505, "August");      // 2024-08-01
    test_case("mmmm", 45536, "September");   // 2024-09-01
    test_case("mmmm", 45566, "October");     // 2024-10-01
    test_case("mmmm", 45597, "November");    // 2024-11-01
    test_case("mmmm", 45627, "December");    // 2024-12-01
    
    // Days of week
    test_case("dddd", 45291, "Sunday");      // 2023-12-31
    test_case("dddd", 45292, "Monday");      // 2024-01-01
    test_case("dddd", 45293, "Tuesday");     // 2024-01-02
    test_case("dddd", 45294, "Wednesday");   // 2024-01-03
    test_case("dddd", 45295, "Thursday");    // 2024-01-04
    test_case("dddd", 45296, "Friday");      // 2024-01-05
    test_case("dddd", 45297, "Saturday");    // 2024-01-06
    
    // ========== Time Formats ==========
    test_case("hh:mm", 0, "00:00");          // Midnight
    test_case("hh:mm", 0.25, "06:00");       // 6 AM
    test_case("hh:mm", 0.5, "12:00");        // Noon
    test_case("hh:mm", 0.75, "18:00");       // 6 PM
    test_case("hh:mm", 1, "00:00");          // Next midnight
    
    test_case("hh:mm:ss", 0, "00:00:00");
    test_case("hh:mm:ss", 0.5, "12:00:00");
    test_case("hh:mm:ss", 0.00001, "00:00:00");  // ~0.86 seconds
    test_case("hh:mm:ss", 0.0000116, "00:00:01"); // ~1 second
    
    test_case("h:mm", 0, "0:00");
    test_case("h:mm", 0.5, "12:00");
    test_case("h:mm", 6.5/24, "6:30");
    
    test_case("h:mm AM/PM", 0, "12:00 AM");
    test_case("h:mm AM/PM", 0.25, "6:00 AM");
    test_case("h:mm AM/PM", 0.5, "12:00 PM");
    test_case("h:mm AM/PM", 0.75, "6:00 PM");
    test_case("h:mm AM/PM", 0.916667, "10:00 PM");  // ~10 PM
    
    test_case("hh:mm:ss.000", 0.0000116, "00:00:01.002");
    
    // ========== Elapsed Time ==========
    test_case("[h]", 0, "0");
    test_case("[h]", 1, "24");
    test_case("[h]", 1.5, "36");
    test_case("[h]", 2, "48");
    test_case("[h]", 10, "240");
    
    test_case("[m]", 0, "0");
    test_case("[m]", 0.5, "720");     // 12 hours
    test_case("[m]", 1, "1440");      // 24 hours
    test_case("[m]", 1.5, "2160");    // 36 hours
    
    test_case("[s]", 0, "0");
    test_case("[s]", 0.5, "43200");   // 12 hours
    test_case("[s]", 1, "86400");     // 24 hours
    test_case("[s]", 1.5, "129600");  // 36 hours
    
    test_case("[h]:mm", 1.5, "36:00");
    test_case("[m]:ss", 1, "1440:00");
    
    // ========== Combined Date-Time ==========
    test_case("yyyy-mm-dd hh:mm", 45292.5, "2024-01-01 12:00");
    test_case("yyyy-mm-dd hh:mm:ss", 45292.5, "2024-01-01 12:00:00");
    test_case("dd/mm/yyyy h:mm AM/PM", 45292.25, "01/01/2024 6:00 AM");
    
    // ========== Multi-Section Formats ==========
    // Positive;Negative
    test_case("0;[Red]-0", 123, "123");
    test_case("0;[Red]-0", -123, "-123");
    
    // Positive;Negative;Zero
    test_case("#,##0;(#,##0);\"Zero\"", 1234, "1,234");
    test_case("#,##0;(#,##0);\"Zero\"", -1234, "(1,234)");
    test_case("#,##0;(#,##0);\"Zero\"", 0, "Zero");
    
    test_case("0.00;-0.00;\"--\"", 1.23, "1.23");
    test_case("0.00;-0.00;\"--\"", -1.23, "-1.23");
    test_case("0.00;-0.00;\"--\"", 0, "--");
    
    // Positive;Negative;Zero;Text
    test_case("0;0;0;\"Text: \"@", 123, "123");
    test_case("0;0;0;\"Text: \"@", -123, "-123");
    test_case("0;0;0;\"Text: \"@", 0, "0");
    test_case("0;0;0;\"Text: \"@", 999, "999");
    
    // ========== Conditional Formats ==========
    test_case("[>100]\"High\";\"Low\"", 150, "High");
    test_case("[>100]\"High\";\"Low\"", 50, "Low");
    test_case("[>100]\"High\";\"Low\"", 100, "Low");  // Not > 100
    
    test_case("[>=100]\"High\";\"Low\"", 100, "High");
    test_case("[>=100]\"High\";\"Low\"", 99, "Low");
    
    test_case("[<100]\"Low\";\"High\"", 50, "Low");
    test_case("[<100]\"Low\";\"High\"", 150, "High");
    test_case("[<100]\"Low\";\"High\"", 100, "High");
    
    test_case("[<=100]\"Low\";\"High\"", 100, "Low");
    test_case("[<=100]\"Low\";\"High\"", 101, "High");
    
    test_case("[=0]\"Zero\";\"Non-Zero\"", 0, "Zero");
    test_case("[=0]\"Zero\";\"Non-Zero\"", 1, "Non-Zero");
    
    test_case("[<>0]\"Non-Zero\";\"Zero\"", 0, "Zero");
    test_case("[<>0]\"Non-Zero\";\"Zero\"", 1, "Non-Zero");
    
    test_case("[>100]\"High\";[<50]\"Low\";\"Medium\"", 150, "High");
    test_case("[>100]\"High\";[<50]\"Low\";\"Medium\"", 25, "Low");
    test_case("[>100]\"High\";[<50]\"Low\";\"Medium\"", 75, "Medium");
    
    // ========== Text Format ==========
    test_text("@", "", "");
    test_text("@", "Hello", "Hello");
    test_text("@", "123", "123");
    test_text("\"Prefix: \"@", "World", "Prefix: World");
    test_text("@\" Suffix\"", "Hello", "Hello Suffix");
    test_text("\"[\"@\"]\"", "Test", "[Test]");
    test_text("\"Name: \"@\"!\"", "Bob", "Name: Bob!");
    
    // ========== Escaped Characters ==========
    test_case("\"Total: \"#,##0", 1234, "Total: 1,234");
    test_case("\"$\"#,##0.00", 1234.56, "$1,234.56");
    test_case("#,##0\" USD\"", 1234, "1,234 USD");
    test_case("\\$#,##0", 1234, "$1,234");
    test_case("\\#\\#0", 123, "##123");
    test_case("0\" units\"", 42, "42 units");
    test_case("\"Value = \"0.00", 1.5, "Value = 1.50");
    
    // ========== Color Codes (parsed, not rendered) ==========
    test_case("[Red]#,##0", 1234, "1,234");
    test_case("[Green]#,##0", 1234, "1,234");
    test_case("[Blue]#,##0.00", 1234.56, "1,234.56");
    test_case("[Black]0", 123, "123");
    test_case("[Cyan]0", 123, "123");
    test_case("[Magenta]0", 123, "123");
    test_case("[Yellow]0", 123, "123");
    test_case("[White]0", 123, "123");
    
    // ========== Complex Formats ==========
    test_case("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)", 1234.56, "$1,234.56 ");  // Accounting format with trailing space from _)
    test_case("_(* #,##0.00_);_(* (#,##0.00)", 1234, "  1,234.00 ");  // Has leading spaces from _( and trailing from _)
    
    // Fraction formats
    test_case("# ?/?", 1.5, "1 1/2");
    test_case("# ?\?/?\?", 1.25, "1 1/4");
    
    // ========== Negative Numbers ==========
    test_case("0", -123, "-123");
    test_case("0.00", -123.456, "-123.46");
    test_case("#,##0", -1234, "-1,234");
    test_case("#,##0.00", -1234.56, "-1,234.56");
    test_case("0%", -0.5, "-50%");
    test_case("0.00E+00", -1234, "-1.23E+03");
    
    // ========== Edge Cases ==========
    test_case("0.00", 0, "0.00");
    test_case("0.00", -0.0, "0.00");
    test_case("0.00", 0.001, "0.00");   // Rounds to 0
    test_case("0.00", 0.005, "0.01");   // Rounds up
    test_case("0.00", -0.005, "-0.01"); // Negative rounds
    test_case("0", 0.4, "0");
    test_case("0", 0.5, "1");           // Round half up
    test_case("0", 0.499999, "0");
    test_case("#,##0.00", 0.001, "0.00");
    test_case("0.00", 999999999.99, "999999999.99");
    test_case("0.00", 0.9999999999, "1.00");
    
    // Very large numbers
    test_case("#,##0", 1234567890123, "1,234,567,890,123");
    test_case("0.00E+00", 1234567890123, "1.23E+12");
    
    // Very small numbers
    test_case("0.00E+00", 0.000000001, "1.00E-09");
    test_case("0.000000000", 0.000000001, "0.000000001");

    // Special values
    test_case("0", 1e-10, "0");
    // Note: "0.0000000001" is not a valid Excel number format (digit '1' is not a placeholder)
    // In Excel, literal digits need to be escaped with backslash or quoted
    test_case("0.000000000\"1\"", 1e-10, "0.0000000001");

    // ========== Summary ==========
    std::cout << std::endl;
    std::cout << "=== Test Summary ===" << std::endl;
    std::cout << "Passed: " << total.passed << std::endl;
    std::cout << "Failed: " << total.failed << std::endl;
    std::cout << "Total:  " << (total.passed + total.failed) << std::endl;
    
    if (total.failed == 0) {
        std::cout << "\n*** ALL TESTS PASSED ***" << std::endl;
    } else {
        std::cout << "\n*** " << total.failed << " TEST(S) FAILED ***" << std::endl;
    }
}

int main()
{
#ifdef _WIN32
    auto __ConOutCP = GetConsoleOutputCP();
    SetConsoleOutputCP(CP_UTF8);
#endif

    test_number_format();

#ifdef _WIN32
    SetConsoleOutputCP(__ConOutCP);
#endif
    return 0;
}
