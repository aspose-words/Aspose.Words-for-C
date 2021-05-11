#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Loading/DocumentDirection.h>
#include <Aspose.Words.Cpp/Loading/TxtLeadingSpacesOptions.h>
#include <Aspose.Words.Cpp/Loading/TxtLoadOptions.h>
#include <Aspose.Words.Cpp/Loading/TxtTrailingSpacesOptions.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Section.h>
#include <system/array.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/io/memory_stream.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;

namespace ApiExamples {

class ExTxtLoadOptions : public ApiExampleBase
{
public:
    void DetectNumberingWithWhitespaces(bool detectNumberingWithWhitespaces)
    {
        //ExStart
        //ExFor:TxtLoadOptions.DetectNumberingWithWhitespaces
        //ExSummary:Shows how to detect lists when loading plaintext documents.
        // Create a plaintext document in a string with four separate parts that we may interpret as lists,
        // with different delimiters. Upon loading the plaintext document into a "Document" object,
        // Aspose.Words will always detect the first three lists and will add a "List" object
        // for each to the document's "Lists" property.
        const String textDoc = String(u"Full stop delimiters:\n") + u"1. First list item 1\n" + u"2. First list item 2\n" + u"3. First list item 3\n\n" +
                               u"Right bracket delimiters:\n" + u"1) Second list item 1\n" + u"2) Second list item 2\n" + u"3) Second list item 3\n\n" +
                               u"Bullet delimiters:\n" + u"• Third list item 1\n" + u"• Third list item 2\n" + u"• Third list item 3\n\n" +
                               u"Whitespace delimiters:\n" + u"1 Fourth list item 1\n" + u"2 Fourth list item 2\n" + u"3 Fourth list item 3";

        // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
        // to modify how we load a plaintext document.
        auto loadOptions = MakeObject<TxtLoadOptions>();

        // Set the "DetectNumberingWithWhitespaces" property to "true" to detect numbered items
        // with whitespace delimiters, such as the fourth list in our document, as lists.
        // This may also falsely detect paragraphs that begin with numbers as lists.
        // Set the "DetectNumberingWithWhitespaces" property to "false"
        // to not create lists from numbered items with whitespace delimiters.
        loadOptions->set_DetectNumberingWithWhitespaces(detectNumberingWithWhitespaces);

        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(textDoc)), loadOptions);

        if (detectNumberingWithWhitespaces)
        {
            ASSERT_EQ(4, doc->get_Lists()->get_Count());
            ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->LINQ_Any(
                [](SharedPtr<Node> p) { return p->GetText().Contains(u"Fourth list") && (System::DynamicCast<Paragraph>(p))->get_IsListItem(); }));
        }
        else
        {
            ASSERT_EQ(3, doc->get_Lists()->get_Count());
            ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->LINQ_Any(
                [](SharedPtr<Node> p) { return p->GetText().Contains(u"Fourth list") && (System::DynamicCast<Paragraph>(p))->get_IsListItem(); }));
        }
        //ExEnd
    }

    void TrailSpaces(TxtLeadingSpacesOptions txtLeadingSpacesOptions, TxtTrailingSpacesOptions txtTrailingSpacesOptions)
    {
        //ExStart
        //ExFor:TxtLoadOptions.TrailingSpacesOptions
        //ExFor:TxtLoadOptions.LeadingSpacesOptions
        //ExFor:TxtTrailingSpacesOptions
        //ExFor:TxtLeadingSpacesOptions
        //ExSummary:Shows how to trim whitespace when loading plaintext documents.
        String textDoc = String(u"      Line 1 \n") + u"    Line 2   \n" + u" Line 3       ";

        // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
        // to modify how we load a plaintext document.
        auto loadOptions = MakeObject<TxtLoadOptions>();

        // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Preserve"
        // to preserve all whitespace characters at the start of every line.
        // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.ConvertToIndent"
        // to remove all whitespace characters from the start of every line,
        // and then apply a left first line indent to the paragraph to simulate the effect of the whitespaces.
        // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Trim"
        // to remove all whitespace characters from every line's start.
        loadOptions->set_LeadingSpacesOptions(txtLeadingSpacesOptions);

        // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Preserve"
        // to preserve all whitespace characters at the end of every line.
        // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Trim" to
        // remove all whitespace characters from the end of every line.
        loadOptions->set_TrailingSpacesOptions(txtTrailingSpacesOptions);

        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(textDoc)), loadOptions);
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        switch (txtLeadingSpacesOptions)
        {
        case TxtLeadingSpacesOptions::ConvertToIndent:
            ASPOSE_ASSERT_EQ(37.8, paragraphs->idx_get(0)->get_ParagraphFormat()->get_FirstLineIndent());
            ASPOSE_ASSERT_EQ(25.2, paragraphs->idx_get(1)->get_ParagraphFormat()->get_FirstLineIndent());
            ASPOSE_ASSERT_EQ(6.3, paragraphs->idx_get(2)->get_ParagraphFormat()->get_FirstLineIndent());
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().StartsWith(u"Line 1"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().StartsWith(u"Line 2"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().StartsWith(u"Line 3"));
            break;

        case TxtLeadingSpacesOptions::Preserve:
            ASSERT_TRUE(paragraphs->LINQ_All([](SharedPtr<Node> p)
                                             { return (System::DynamicCast<Paragraph>(p))->get_ParagraphFormat()->get_FirstLineIndent() == 0.0; }));
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().StartsWith(u"      Line 1"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().StartsWith(u"    Line 2"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().StartsWith(u" Line 3"));
            break;

        case TxtLeadingSpacesOptions::Trim:
            ASSERT_TRUE(paragraphs->LINQ_All([](SharedPtr<Node> p)
                                             { return (System::DynamicCast<Paragraph>(p))->get_ParagraphFormat()->get_FirstLineIndent() == 0.0; }));
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().StartsWith(u"Line 1"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().StartsWith(u"Line 2"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().StartsWith(u"Line 3"));
            break;
        }

        switch (txtTrailingSpacesOptions)
        {
        case TxtTrailingSpacesOptions::Preserve:
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().EndsWith(u"Line 1 \r"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().EndsWith(u"Line 2   \r"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().EndsWith(u"Line 3       \f"));
            break;

        case TxtTrailingSpacesOptions::Trim:
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().EndsWith(u"Line 1\r"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().EndsWith(u"Line 2\r"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().EndsWith(u"Line 3\f"));
            break;
        }
        //ExEnd
    }

    void DetectDocumentDirection()
    {
        //ExStart
        //ExFor:TxtLoadOptions.DocumentDirection
        //ExFor:ParagraphFormat.Bidi
        //ExSummary:Shows how to detect plaintext document text direction.
        // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
        // to modify how we load a plaintext document.
        auto loadOptions = MakeObject<TxtLoadOptions>();

        // Set the "DocumentDirection" property to "DocumentDirection.Auto" automatically detects
        // the direction of every paragraph of text that Aspose.Words loads from plaintext.
        // Each paragraph's "Bidi" property will store its direction.
        loadOptions->set_DocumentDirection(DocumentDirection::Auto);

        // Detect Hebrew text as right-to-left.
        auto doc = MakeObject<Document>(MyDir + u"Hebrew text.txt", loadOptions);

        ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Bidi());

        // Detect English text as right-to-left.
        doc = MakeObject<Document>(MyDir + u"English text.txt", loadOptions);

        ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Bidi());
        //ExEnd
    }
};

} // namespace ApiExamples
