#include <xlsxtext.hpp>

#include <iostream>
#include <cstring>
#include <ctime>

int main()
{
    clock_t start = clock();
    xlsxtext::workbook workbook("../doc/zip.xlsx");
    workbook.read();
    for (auto worksheet : workbook)
    {
        std::cout << worksheet.name() << std::endl;
        try
        {
            auto errors = worksheet.read();
            for (auto [refer, msg] : errors)
                std::cerr << refer.value() << ": " << msg << std::endl;

            std::cout << std::endl;
            for (auto row : worksheet)
            {
                for (auto cell : row)
                {
                    std::cout << cell.refer.value() << ": " << cell.value << std::endl;
                }
                std::cout << std::endl;
            }
        }
        catch (std::string err)
        {
            std::cout << err << std::endl;
        }
    }
    std::cout << clock() - start << std::endl;

    return 0;
}
