#pragma once

// #define USE_PUGIXML

#include "miniz/miniz.h"
#ifdef USE_PUGIXML
#include "pugixml/pugixml.hpp"
#endif

#include <cstring>
#include <string>
#include <vector>
#include <map>
#include <memory>

namespace xlsxtext
{
#ifndef USE_PUGIXML
    namespace internal
    {
        class xml_reader
        {
        public:
            enum class node_type
            {
                null,
                error,
                element_start,
                element_end,
                element,
                text,
            };
            static std::string text(std::string_view value) noexcept
            {
                std::string text = std::string(value);
                for (std::string::size_type i = 0; i < text.size(); ++i)
                {
                    if (text[i] == '&')
                    {
                        if (i <= text.size() - 6 && text[i + 1] == 'a' && text[i + 2] == 'p' && text[i + 3] == 'o' && text[i + 4] == 's' && text[i + 5] == ';')
                            text.replace(i, 6, "'");
                        else if (i <= text.size() - 5 && text[i + 1] == 'a' && text[i + 2] == 'm' && text[i + 3] == 'p' && text[i + 4] == ';')
                            text.replace(i, 5, "&");
                        else if (i <= text.size() - 6 && text[i + 1] == 'q' && text[i + 2] == 'u' && text[i + 3] == 'o' && text[i + 4] == 't' && text[i + 5] == ';')
                            text.replace(i, 6, "\"");
                        else if (i <= text.size() - 4 && text[i + 1] == 'l' && text[i + 2] == 't' && text[i + 3] == ';')
                            text.replace(i, 4, "<");
                        else if (i <= text.size() - 4 && text[i + 1] == 'g' && text[i + 2] == 't' && text[i + 3] == ';')
                            text.replace(i, 4, ">");
                        else if (i <= text.size() - 4 && text[i + 1] == '#')
                        {
                            bool ok = true;
                            unsigned code = 0, base = 10;
                            auto pos = i + 2;
                            if (text[pos] == 'x')
                            {
                                base = 16;
                                ++pos;
                            }
                            for (auto ch = text[pos]; pos < text.size() && ch != ';'; ch = text[++pos])
                            {
                                code *= base;
                                if (ch >= '0' && ch <= '9')
                                    code += ch - '0';
                                else if (base == 16 && ch - 'A' <= 5)
                                    code += ch - 'A' + 10;
                                else if (base == 16 && ch - 'a' <= 5)
                                    code += ch - 'a' + 10;
                                else
                                {
                                    ok = false;
                                    break;
                                }
                            }
                            if (ok && pos < text.size() && text[pos] == ';')
                            {
                                char str[4] = {};
                                if (code < 0x80)
                                {
                                    str[0] = static_cast<uint8_t>(code);
                                }
                                else if (code < 0x800)
                                {
                                    str[0] = static_cast<uint8_t>(0xC0 | (code >> 6));
                                    str[1] = static_cast<uint8_t>(0x80 | (code & 0x3F));
                                }
                                else if (code <= 0xFFFF)
                                {
                                    str[0] = static_cast<uint8_t>(0xE0 | (code >> 12));
                                    str[1] = static_cast<uint8_t>(0x80 | ((code >> 6) & 0x3F));
                                    str[2] = static_cast<uint8_t>(0x80 | (code & 0x3F));
                                }
                                else
                                {
                                    str[0] = static_cast<uint8_t>(0xF0 | (code >> 18));
                                    str[1] = static_cast<uint8_t>(0x80 | ((code >> 12) & 0x3F));
                                    str[2] = static_cast<uint8_t>(0x80 | ((code >> 6) & 0x3F));
                                    str[3] = static_cast<uint8_t>(0x80 | (code & 0x3F));
                                }
                                text.replace(i, pos - i + 1, str);
                            }
                        }
                    }
                }
                return text;
            }

        private:
            std::string_view _data;
            std::string_view::size_type _pos;

        public:
            xml_reader(std::string_view data, std::string_view::size_type pos = 0) noexcept : _data(data), _pos(pos) {}

        private:
            node_type _type = node_type::null;
            std::string_view _name;
            std::map<std::string_view, std::string_view> _attributes;
            std::string_view _value;

            std::vector<std::string_view> _tree;

            void reset() noexcept
            {
                _type = node_type::null;
                _name = _value = _data.substr(0, 0);
                _attributes.clear();
            }

            bool is_name_char(unsigned char c, bool first = false) const noexcept { return ('A' <= c && c <= 'Z') || ('a' <= c && c <= 'z') || c == '_' || c > 127 || (!first && (('0' <= c && c <= '9') || c == ':' || c == '-' || c == '.')); }
            bool is_space_char(unsigned char c) const noexcept { return c == ' ' || c == '\t' || c == '\r' || c == '\n'; }

        public:
            node_type type() const noexcept { return _type; }
            std::string_view name() const noexcept { return _name; }
            std::map<std::string_view, std::string_view> attributes() const noexcept { return _attributes; }
            std::string_view value() const noexcept { return _value; }

            bool read() noexcept
            {
                if (_type == node_type::error)
                    return false;

                if (_type == node_type::element_start)
                    _tree.push_back(_name);

                reset();

                auto size = _data.size();
                while (_pos < size)
                {
                    auto pos = _pos;
                    while (pos < size && is_space_char(_data[pos]))
                        ++pos;
                    if (pos >= size)
                    {
                        _pos = pos;
                        return false;
                    }
                    if (_data[pos] != '<')
                    {
                        _type = node_type::text;
                        while (pos < size && _data[pos] != '<')
                            ++pos;
                        _value = _data.substr(_pos, pos - _pos);
                        _pos = pos;
                        return true;
                    }
                    else
                    {
                        /**
                         * <?xxx?>
                         * <!x>
                         * <![CDATA["xxxxx"]]>
                         * <!--xxx-->
                         */
                        _type = node_type::error;
                        _pos = pos++;

                        if (pos > size - 3) // <x/>
                            return false;
                        if (_data[pos] == '?')
                        {
                            ++pos;
                            char quote = 0;
                            while (pos < size - 1 && (quote || _data[pos] != '?' || _data[pos + 1] != '>'))
                            {
                                if (_data[pos] == '"' || _data[pos] == '\'')
                                    quote = quote == 0 ? _data[pos] : quote == _data[pos] ? 0 : quote;
                                ++pos;
                            }
                            if (quote || pos > size - 2)
                                return false;
                            _pos = pos + 2;
                            continue;
                        }
                        else if (_data[pos] == '!')
                        {
                            if (_data[pos + 1] == '[')
                            {
                                pos += 2;
                                while (pos < size - 2 && (_data[pos] != ']' || _data[pos + 1] != ']' || _data[pos + 2] != '>'))
                                    ++pos;
                                if (pos > size - 3)
                                    return false;
                                _pos = pos + 3;
                                continue;
                            }
                            else if (_data[pos + 1] == '-' && _data[pos + 2] == '-')
                            {
                                pos += 3;
                                while (pos < size - 2 && (_data[pos] != '-' || _data[pos + 1] != '-' || _data[pos + 2] != '>'))
                                    ++pos;
                                if (pos > size - 3)
                                    return false;
                                _pos = pos + 3;
                                continue;
                            }
                            else
                            {
                                ++pos;
                                char quote = 0;
                                while (pos < size && (quote || _data[pos] != '>'))
                                {
                                    if (_data[pos] == '"' || _data[pos] == '\'')
                                        quote = quote == 0 ? _data[pos] : quote == _data[pos] ? 0 : quote;
                                    ++pos;
                                }
                                if (pos >= size)
                                    return false;
                                _pos = pos + 1;
                                continue;
                            }
                        }

                        if (is_name_char(_data[pos], true))
                        {
                            while (pos < size && is_name_char(_data[pos])) // skip name
                                ++pos;
                            if (pos > size - 2 || (!is_space_char(_data[pos]) && _data[pos] != '/' && _data[pos] != '>'))
                                return false;
                            decltype(_name) name = _data.substr(_pos + 1, pos - _pos - 1);

                            decltype(_attributes) attributes;
                            while (pos < size) // read attributes
                            {
                                while (pos < size && is_space_char(_data[pos])) // skip space
                                    ++pos;
                                if (pos > size - 2 || (!is_name_char(_data[pos], true) && _data[pos] != '/' && _data[pos] != '>'))
                                    return false;

                                if (_data[pos] == '/' || _data[pos] == '>')
                                    break;

                                decltype(size) key_pos = pos++, key_size = 0;
                                while (pos < size && is_name_char(_data[pos])) // skip name
                                    ++pos;
                                if (pos > size - 2)
                                    return false;
                                key_size = pos - key_pos;

                                while (pos < size && is_space_char(_data[pos])) // skip space
                                    ++pos;
                                if (pos > size - 5 || _data[pos] != '=') //=""/> or =''/>
                                    return false;

                                ++pos;
                                while (pos < size && is_space_char(_data[pos])) // skip space
                                    ++pos;
                                if (pos > size - 4 || (_data[pos] != '\'' && _data[pos] != '\"')) //""/> or ''/>
                                    return false;

                                ++pos;
                                auto sign = _data[pos - 1];
                                decltype(size) value_pos = pos, value_size = 0;
                                while (pos < size && _data[pos] != sign) // "/> or '/>
                                    ++pos;
                                if (pos > size - 3 || _data[pos] != sign)
                                    return false;
                                value_size = pos - value_pos;

                                ++pos;
                                if (!is_space_char(_data[pos]) && _data[pos] != '/' && _data[pos] != '>')
                                    return false;

                                attributes[_data.substr(key_pos, key_size)] = _data.substr(value_pos, value_size);
                            }
                            if (pos > size - 2 || (_data[pos] != '>' && (_data[pos] != '/' || _data[pos + 1] != '>'))) // not end with '>' or '/>'
                                return false;

                            if (_data[pos] == '>')
                            {
                                _type = node_type::element_start;
                                _pos = pos + 1;
                            }
                            else
                            {
                                _type = node_type::element;
                                _pos = pos + 2;
                            }
                            _name = name;
                            _attributes = attributes;
                            return true;
                        }
                        else if (_data[pos] == '/')
                        {
                            if (!is_name_char(_data[pos + 1], true))
                            {
                                _type = node_type::error;
                                return false;
                            }

                            pos += 2;
                            while (pos < size && is_name_char(_data[pos]))
                                ++pos;
                            if (pos >= size)
                                return false;

                            auto name = _data.substr(_pos + 2, pos - _pos - 2);
                            if (_tree.size() <= 0 || _tree[_tree.size() - 1] != name)
                                return false;

                            while (pos < size && is_space_char(_data[pos]))
                                ++pos;
                            if (pos >= size || _data[pos] != '>')
                                return false;

                            _tree.pop_back();
                            _pos = pos + 1;
                            _type = node_type::element_end;
                            _name = name;
                            return true;
                        }
                    }
                }
                return false;
            }
        };
    } // namespace internal
#endif

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

            virtual ~package() { mz_zip_reader_end(&archive); }
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
        std::string read_value(const reference &refer, const std::string &v, const std::string &t, const std::string &s, std::vector<std::pair<reference, std::string>> &errors) const
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
                errors.push_back({refer, "date type is not supported"});
            }
            else if (t == "e")
            {
                errors.push_back({refer, "date type is not supported"});
            }
            else
            {
                if (s == "")
                    return v;

                auto index = std::stol(s);
                auto xf = index < _package->cell_xfs.size() ? _package->cell_xfs[index] : 0;
                const auto &code = _package->numfmts.find(xf) != _package->numfmts.end() ? _package->numfmts[xf] : "";
                if (code == "General" || code == "@")
                    return v;
                else
                    errors.push_back({refer, "the numFmt code:\"" + code + "\"is not supported"});
            }
            return "";
        }

    public:
        std::string name() const noexcept { return _name; }
        std::vector<std::pair<reference, std::string>> read()
        {
            _merge_cells.clear();
            _rows.clear();

            std::vector<std::pair<reference, std::string>> errors;

            void *buffer = nullptr;
            size_t size = 0;
            if (buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, _part.c_str(), &size, 0))
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
#ifdef USE_PUGIXML
                pugi::xml_document doc;
                auto result = doc.load_buffer(buffer, size);
                mz_free(buffer);
                if (!result)
                {
                    errors.push_back({_name, "open failed"});
                    return errors;
                }

                for (auto mc = doc.child("worksheet").child("mergeCells").child("mergeCell"); mc; mc = mc.next_sibling("mergeCell"))
                {
                    if (auto ref = mc.attribute("ref"); ref)
                    {
                        std::string refs = ref.value();
                        if (auto split = refs.find(':'); split != std::string::npos && split < refs.size() - 1)
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
                        reference refer(row.attribute("r").value()); // "r" is optional
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
                            if (auto r = is.child("r"); r)
                            {
                                for (; r; r = r.next_sibling("r"))
                                    v += r.child("t").text().get();
                            }
                            else
                                v = is.child("t").text().get();
                        }

                        std::string value = read_value(refer, v, t, s, errors);
                        if (value != "")
                            cells.push_back(cell(refer, value));
                    }

                    if (cells.size())
                        _rows.push_back(cells);
                }
#else
                using internal::xml_reader;

                xml_reader reader(std::string_view((const char *)buffer, size));
                char path[6] = {}, depth = 0;

                std::vector<cell> cells;
                std::string r, t, s, v, iv;
                while (reader.read())
                {
                    auto type = reader.type();
                    if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start || type == xml_reader::node_type::element_end)
                    {
                        auto name = reader.name();
                        if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start)
                        {
                            if (depth == 2 && path[0] == 1 && path[1] == 2 && name == "mergeCell")
                            {
                                if (auto refs = reader.attributes()["ref"]; refs != "")
                                {
                                    if (auto split = refs.find(':'); split != std::string::npos && split < refs.size() - 1)
                                        _merge_cells.push_back({std::string(refs.substr(0, split)), std::string(refs.substr(split + 1)), ""});
                                }
                            }
                            else if (depth == 3 && path[0] == 1 && path[1] == 1 && path[2] == 1 && name == "c")
                            {
                                auto attributes = reader.attributes();
                                r = attributes["r"], t = attributes["t"], s = attributes["s"];
                            }
                        }
                        if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_end)
                        {
                            if (depth == 4 && path[0] == 1 && path[1] == 1 && path[2] == 1 && path[3] == 1 && name == "c")
                            {
                                std::string value = read_value(r, v, t, s, errors);
                                cells.push_back(cell(r, value));
                                r = t = s = v = iv = "";
                            }
                            if (depth == 3 && path[0] == 1 && path[1] == 1 && path[2] == 1 && name == "row")
                            {
                                if (cells.size() > 0)
                                    _rows.push_back(cells);
                                cells.clear();
                            }
                        }

                        if (type == xml_reader::node_type::element_start)
                        {
                            if (depth == 0)
                                path[depth] = name == "worksheet" ? 1 : 0;
                            else if (depth == 1)
                                path[depth] = name == "sheetData" ? 1 : name == "mergeCells" ? 2 : 0;
                            else if (depth == 2)
                                path[depth] = name == "row" ? 1 : 0;
                            else if (depth == 3)
                                path[depth] = name == "c" ? 1 : 0;
                            else if (depth == 4)
                                path[depth] = name == "v" ? 1 : name == "is" ? 2 : 0;
                            else if (depth == 5)
                                path[depth] = name == "t" ? 1 : 0;
                            ++depth;
                        }
                        else if (type == xml_reader::node_type::element_end)
                        {
                            --depth;
                        }
                    }
                    else if (type == xml_reader::node_type::text)
                    {
                        if (path[0] == 1 && path[1] == 1 && path[2] == 1 && path[3] == 1)
                        {
                            if (depth == 5 && path[4] == 1)
                                v = xml_reader::text(reader.value());
                            else if (depth == 6 && path[4] == 2 && path[5] == 1)
                                iv += xml_reader::text(reader.value());
                        }
                    }
                }
                mz_free(buffer);
#endif
                return errors;
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

#ifdef USE_PUGIXML
            pugi::xml_document doc;
#else
            using internal::xml_reader;
#endif

            if (buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, "_rels/.rels", &size, 0))
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
#ifdef USE_PUGIXML
                auto result = doc.load_buffer(buffer, size);
                mz_free(buffer);
                if (!result)
                    return false;

                for (auto res = doc.child("Relationships").child("Relationship"); res; res = res.next_sibling("Relationship"))
                {
                    if (auto type = res.attribute("Type"), target = res.attribute("Target");
                        type && target && !std::strcmp(type.value(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"))
                    {
                        workbook_part = target.value();
                        break;
                    }
                }
#else
                xml_reader reader(std::string_view((const char *)buffer, size));
                char path[1] = {}, depth = 0;

                while (reader.read())
                {
                    auto type = reader.type();
                    if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start || type == xml_reader::node_type::element_end)
                    {
                        auto name = reader.name();
                        if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start)
                        {
                            if (depth == 1 && path[0] == 1 && name == "Relationship")
                            {
                                auto attributes = reader.attributes();
                                if (auto type = attributes["Type"], target = attributes["Target"];
                                    target != "" && type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
                                {
                                    workbook_part = target;
                                    break;
                                }
                            }
                        }

                        if (type == xml_reader::node_type::element_start)
                        {
                            if (depth == 0)
                                path[depth] = name == "Relationships" ? 1 : 0;
                            ++depth;
                        }
                        else if (type == xml_reader::node_type::element_end)
                        {
                            --depth;
                        }
                    }
                }
                mz_free(buffer);
#endif
            }
            if (buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, "xl/_rels/workbook.xml.rels", &size, 0))
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
#ifdef USE_PUGIXML
                auto result = doc.load_buffer(buffer, size);
                mz_free(buffer);
                if (!result)
                    return false;

                shared_strings_part = styles_part = "";
                for (auto res = doc.child("Relationships").child("Relationship"); res; res = res.next_sibling("Relationship"))
                {
                    if (auto id = res.attribute("Id"), type = res.attribute("Type"), target = res.attribute("Target");
                        id && type && target)
                    {
                        if (!std::strcmp(type.value(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"))
                            shared_strings_part = std::string("xl/") + target.value();
                        else if (!std::strcmp(type.value(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"))
                            styles_part = std::string("xl/") + target.value();
                        else if (!std::strcmp(type.value(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"))
                            sheets[id.value()] = std::string("xl/") + target.value();
                    }
                }
#else
                xml_reader reader(std::string_view((const char *)buffer, size));
                char path[1] = {}, depth = 0;

                shared_strings_part = styles_part = "";
                while (reader.read())
                {
                    auto type = reader.type();
                    if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start || type == xml_reader::node_type::element_end)
                    {
                        auto name = reader.name();
                        if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start)
                        {
                            if (depth == 1 && path[0] == 1 && name == "Relationship")
                            {
                                auto attributes = reader.attributes();
                                if (auto id = attributes["Id"], type = attributes["Type"], target = attributes["Target"]; id != "" && target != "")
                                {
                                    if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
                                        shared_strings_part = "xl/" + std::string(target);
                                    else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
                                        styles_part = "xl/" + std::string(target);
                                    else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                                        sheets[std::string(id)] = "xl/" + std::string(target);
                                }
                            }
                        }

                        if (type == xml_reader::node_type::element_start)
                        {
                            if (depth == 0)
                                path[depth] = name == "Relationships" ? 1 : 0;
                            ++depth;
                        }
                        else if (type == xml_reader::node_type::element_end)
                        {
                            --depth;
                        }
                    }
                }
                mz_free(buffer);
#endif
            }
            if (shared_strings_part != "" && (buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, shared_strings_part.c_str(), &size, 0)))
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
#ifdef USE_PUGIXML
                auto result = doc.load_buffer(buffer, size);
                mz_free(buffer);
                if (!result)
                    return false;

                for (auto si = doc.child("sst").child("si"); si; si = si.next_sibling("si"))
                {
                    std::string t;
                    if (auto r = si.child("r"); r)
                    {
                        for (; r; r = r.next_sibling("r"))
                            t += r.child("t").text().get();
                    }
                    else
                        t = si.child("t").text().get();
                    _package->shared_strings.push_back(t);
                }
#else
                xml_reader reader(std::string_view((const char *)buffer, size));
                char path[4] = {}, depth = 0;

                std::string text;
                while (reader.read())
                {
                    auto type = reader.type();
                    if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start || type == xml_reader::node_type::element_end)
                    {
                        auto name = reader.name();
                        if (type == xml_reader::node_type::element_start)
                        {
                            if (depth == 0)
                                path[depth] = name == "sst" ? 1 : 0;
                            else if (depth == 1)
                                path[depth] = name == "si" ? 1 : 0;
                            else if (depth == 2)
                                path[depth] = name == "t" ? 1 : name == "r" ? 2 : 0;
                            else if (depth == 3)
                                path[depth] = name == "t" ? 1 : 0;
                            ++depth;
                        }
                        else if (type == xml_reader::node_type::element_end)
                        {
                            if (depth == 2 && path[0] == 1 && path[1] == 1 && name == "si")
                            {
                                _package->shared_strings.push_back(text);
                                text = "";
                            }
                            --depth;
                        }
                    }
                    else if (type == xml_reader::node_type::text)
                    {
                        if (depth == 3 && path[0] == 1 && path[1] == 1 && path[2] == 1)
                            text = xml_reader::text(reader.value());
                        else if (depth == 4 && path[0] == 1 && path[1] == 1 && path[2] == 2 && path[3] == 1)
                            text += xml_reader::text(reader.value());
                    }
                }
                mz_free(buffer);
#endif
            }
            if (styles_part != "" && (buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, styles_part.c_str(), &size, 0)))
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
#ifdef USE_PUGIXML
                auto result = doc.load_buffer(buffer, size);
                mz_free(buffer);
                if (!result)
                    return false;

                if (auto style = doc.child("styleSheet"); style)
                {
                    for (auto nf = style.child("numFmts").child("numFmt"); nf; nf = nf.next_sibling("numFmt"))
                    {
                        if (auto id = nf.attribute("numFmtId"), code = nf.attribute("formatCode"); id && code)
                            _package->numfmts[std::stol(id.value())] = code.value();
                    }
                    for (auto xf = style.child("cellXfs").child("xf"); xf; xf = xf.next_sibling("xf"))
                    {
                        if (auto id = xf.attribute("numFmtId"); id)
                            _package->cell_xfs.push_back(std::stol(id.value()));
                    }
                }
#else
                xml_reader reader(std::string_view((const char *)buffer, size));
                char path[2] = {}, depth = 0;

                while (reader.read())
                {
                    auto type = reader.type();
                    if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start || type == xml_reader::node_type::element_end)
                    {
                        auto name = reader.name();
                        if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start)
                        {
                            if (depth == 2 && path[0] == 1)
                            {
                                if (path[1] == 1 && name == "numFmt")
                                {
                                    auto attributes = reader.attributes();
                                    if (auto id = attributes["numFmtId"], code = attributes["formatCode"]; id != "" && code != "")
                                        _package->numfmts[std::stol(std::string(id))] = xml_reader::text(code);
                                }
                                else if (path[1] == 2 && name == "xf")
                                {
                                    auto attributes = reader.attributes();
                                    if (auto id = attributes["numFmtId"]; id != "")
                                        _package->cell_xfs.push_back(std::stol(std::string(id)));
                                }
                            }
                        }

                        if (type == xml_reader::node_type::element_start)
                        {
                            if (depth == 0)
                                path[depth] = name == "styleSheet" ? 1 : 0;
                            else if (depth == 1)
                                path[depth] = name == "numFmts" ? 1 : name == "cellXfs" ? 2 : 0;
                            ++depth;
                        }
                        else if (type == xml_reader::node_type::element_end)
                        {
                            --depth;
                        }
                    }
                }
                mz_free(buffer);
#endif
            }
            if (buffer = mz_zip_reader_extract_file_to_heap(&_package->archive, workbook_part.c_str(), &size, 0))
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
                 * <xsd:complexType name="CT_Workbook">
                 *     <xsd:sequence>
                 *         <xsd:element name="sheets" type="CT_Sheets" minOccurs="1" maxOccurs="1"/>
                 *     </xsd:sequence>
                 * </xsd:complexType>
                 * <xsd:element name="workbook" type="CT_Workbook"/>
                 * 
                 * <workbook>
                 *     <sheets>
                 *         <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                 *         <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
                 *     </sheets>
                 * </workbook>
                 */
#ifdef USE_PUGIXML
                auto result = doc.load_buffer(buffer, size);
                mz_free(buffer);
                if (!result)
                    return false;

                for (auto sheet = doc.child("workbook").child("sheets").child("sheet"); sheet; sheet = sheet.next_sibling("sheet"))
                {
                    if (auto name = sheet.attribute("name"), rid = sheet.attribute("r:id"); name && rid && sheets.find(rid.value()) != sheets.end())
                    {
                        const auto &part = sheets[rid.value()];
                        if (mz_zip_reader_locate_file(&_package->archive, part.c_str(), nullptr, 0) != -1)
                            _worksheets.push_back(worksheet::create(name.value(), part, _package));
                    }
                }
#else
                xml_reader reader(std::string_view((const char *)buffer, size));
                char path[2] = {}, depth = 0;

                while (reader.read())
                {
                    auto type = reader.type();
                    if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start || type == xml_reader::node_type::element_end)
                    {
                        auto name = reader.name();
                        if (type == xml_reader::node_type::element || type == xml_reader::node_type::element_start)
                        {
                            if (depth == 2 && path[0] == 1 && path[1] == 1 && name == "sheet")
                            {
                                auto attributes = reader.attributes();
                                if (auto name = std::string(attributes["name"]), rid = std::string(attributes["r:id"]);
                                    name != "" && rid != "" && sheets.find(rid) != sheets.end())
                                {
                                    const auto &part = sheets[rid];
                                    if (mz_zip_reader_locate_file(&_package->archive, part.c_str(), nullptr, 0) != -1)
                                        _worksheets.push_back(worksheet::create(name, part, _package));
                                }
                            }
                        }

                        if (type == xml_reader::node_type::element_start)
                        {
                            if (depth == 0)
                                path[depth] = name == "workbook" ? 1 : 0;
                            else if (depth == 1)
                                path[depth] = name == "sheets" ? 1 : 0;
                            ++depth;
                        }
                        else if (type == xml_reader::node_type::element_end)
                        {
                            --depth;
                        }
                    }
                }
                mz_free(buffer);
#endif
            }

            return true;
        }

        const std::vector<worksheet> &worksheets() const noexcept { return _worksheets; }
        std::vector<worksheet>::const_iterator begin() const noexcept { return _worksheets.begin(); }
        std::vector<worksheet>::const_iterator end() const noexcept { return _worksheets.end(); }
    };
} // namespace xlsxtext