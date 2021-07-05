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
                pugi::xml_document doc;
                auto result = doc.load_buffer_inplace_own(buffer, size);
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
                        _rows.push_back(std::move(cells));
                }
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

            pugi::xml_document doc;
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
                auto result = doc.load_buffer_inplace_own(buffer, size);
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
                auto result = doc.load_buffer_inplace_own(buffer, size);
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
                auto result = doc.load_buffer_inplace_own(buffer, size);
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
                auto result = doc.load_buffer_inplace_own(buffer, size);
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
                auto result = doc.load_buffer_inplace_own(buffer, size);
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
            }

            return true;
        }

        const std::vector<worksheet> &worksheets() const noexcept { return _worksheets; }
        std::vector<worksheet>::const_iterator begin() const noexcept { return _worksheets.begin(); }
        std::vector<worksheet>::const_iterator end() const noexcept { return _worksheets.end(); }
    };
} // namespace xlsxtext