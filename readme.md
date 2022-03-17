**Example**
```
    xlsxtext::workbook workbook("../doc/zip.xlsx");
    workbook.read();
    for (auto worksheet : workbook)
    {
        std::cout << worksheet.name() << std::endl;
        try
        {
            auto errors = worksheet.read();
            for (auto [refer, msg] : errors)
                std::cerr << refer << ": " << msg << std::endl;

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
```

**Thanks**
- pugixml: https://github.com/zeux/pugixml.git
- miniz:https://github.com/richgel999/miniz.git
