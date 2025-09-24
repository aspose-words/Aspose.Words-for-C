#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExPageSetup : public ApiExampleBase
{
    typedef ExPageSetup ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void ClearFormatting();
    void DifferentFirstPageHeaderFooter(bool differentFirstPageHeaderFooter);
    void OddAndEvenPagesHeaderFooter(bool oddAndEvenPagesHeaderFooter);
    void CharactersPerLine();
    void LinesPerPage();
    void SetSectionStart();
    void PageMargins();
    void PaperSizes();
    void ColumnsSameWidth();
    void CustomColumnWidth();
    void VerticalLineBetweenColumns(bool lineBetween);
    void LineNumbers();
    void PageBorderProperties();
    void PageBorders();
    void PageNumbering();
    void FootnoteOptions();
    void Bidi(bool reverseColumns);
    void PageBorder();
    void Gutter();
    void Booklet();
    void SetTextOrientation();
    void SuppressEndnotes();
    void ChapterPageSeparator();
    void JisbPaperSize();
    
protected:

    /// <summary>
    /// Append a section with text and an endnote to a document.
    /// </summary>
    static void InsertSectionWithEndnote(System::SharedPtr<Aspose::Words::Document> doc, System::String sectionBodyText, System::String endnoteText);
    static void TestSuppressEndnotes(System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


