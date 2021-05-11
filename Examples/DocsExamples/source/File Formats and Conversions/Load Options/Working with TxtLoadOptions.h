#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Loading/DocumentDirection.h>
#include <Aspose.Words.Cpp/Loading/TxtLeadingSpacesOptions.h>
#include <Aspose.Words.Cpp/Loading/TxtLoadOptions.h>
#include <Aspose.Words.Cpp/Loading/TxtTrailingSpacesOptions.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <system/array.h>
#include <system/convert.h>
#include <system/io/memory_stream.h>
#include <system/text/encoding.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Load_Options {

class WorkingWithTxtLoadOptions : public DocsExamplesBase
{
public:
    void DetectNumberingWithWhitespaces()
    {
        //ExStart:DetectNumberingWithWhitespaces
        // Create a plaintext document in the form of a string with parts that may be interpreted as lists.
        // Upon loading, the first three lists will always be detected by Aspose.Words,
        // and List objects will be created for them after loading.
        const String textDoc = String(u"Full stop delimiters:\n") + u"1. First list item 1\n" + u"2. First list item 2\n" + u"3. First list item 3\n\n" +
                               u"Right bracket delimiters:\n" + u"1) Second list item 1\n" + u"2) Second list item 2\n" + u"3) Second list item 3\n\n" +
                               u"Bullet delimiters:\n" + u"• Third list item 1\n" + u"• Third list item 2\n" + u"• Third list item 3\n\n" +
                               u"Whitespace delimiters:\n" + u"1 Fourth list item 1\n" + u"2 Fourth list item 2\n" + u"3 Fourth list item 3";

        // The fourth list, with whitespace inbetween the list number and list item contents,
        // will only be detected as a list if "DetectNumberingWithWhitespaces" in a LoadOptions object is set to true,
        // to avoid paragraphs that start with numbers being mistakenly detected as lists.
        auto loadOptions = MakeObject<TxtLoadOptions>();
        loadOptions->set_DetectNumberingWithWhitespaces(true);

        // Load the document while applying LoadOptions as a parameter and verify the result.
        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(textDoc)), loadOptions);

        doc->Save(ArtifactsDir + u"WorkingWithTxtLoadOptions.DetectNumberingWithWhitespaces.docx");
        //ExEnd:DetectNumberingWithWhitespaces
    }

    void HandleSpacesOptions()
    {
        //ExStart:HandleSpacesOptions
        const String textDoc = String(u"      Line 1 \n") + u"    Line 2   \n" + u" Line 3       ";

        auto loadOptions = MakeObject<TxtLoadOptions>();
        loadOptions->set_LeadingSpacesOptions(TxtLeadingSpacesOptions::Trim);
        loadOptions->set_TrailingSpacesOptions(TxtTrailingSpacesOptions::Trim);

        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(textDoc)), loadOptions);

        doc->Save(ArtifactsDir + u"WorkingWithTxtLoadOptions.HandleSpacesOptions.docx");
        //ExEnd:HandleSpacesOptions
    }

    void DocumentTextDirection()
    {
        //ExStart:DocumentTextDirection
        auto loadOptions = MakeObject<TxtLoadOptions>();
        loadOptions->set_DocumentDirection(DocumentDirection::Auto);

        auto doc = MakeObject<Document>(MyDir + u"Hebrew text.txt", loadOptions);

        SharedPtr<Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
        std::cout << System::Convert::ToString(paragraph->get_ParagraphFormat()->get_Bidi()) << std::endl;

        doc->Save(ArtifactsDir + u"WorkingWithTxtLoadOptions.DocumentTextDirection.docx");
        //ExEnd:DocumentTextDirection
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Load_Options
