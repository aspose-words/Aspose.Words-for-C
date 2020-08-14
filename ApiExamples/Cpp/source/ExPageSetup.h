#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExPageSetup : public ApiExampleBase
{
public:

    void ClearFormatting();
    void DifferentHeaders();
    void SetSectionStart();
    void PageMargins();
    void ColumnsSameWidth();
    void CustomColumnWidth();
    void LineNumbers();
    void PageBorderProperties();
    void PageBorders();
    void PageNumbering();
    void FootnoteOptions();
    void Bidi();
    void PageBorder();
    void Gutter();
    void Booklet();
    void SectionTextOrientation();
    void SuppressEndnotes();
    
protected:

    /// <summary>
    /// Add a section to the end of a document, give it a body and a paragraph, then add text and an endnote to that paragraph.
    /// </summary>
    static void InsertSection(System::SharedPtr<Aspose::Words::Document> doc, System::String sectionBodyText, System::String endnoteText);
    static void TestSuppressEndnotes(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


