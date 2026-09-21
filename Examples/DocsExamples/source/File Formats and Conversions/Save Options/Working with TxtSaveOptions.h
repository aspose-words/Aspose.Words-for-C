#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/TxtListIndentation.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/TxtExportHeadersFootersMode.h>
#include <Aspose.Words.Cpp/Saving/TxtSaveOptions.h>
#include <Aspose.Words.Cpp/Section.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithTxtSaveOptions : public DocsExamplesBase
{
public:
    void AddBidiMarks()
    {
        //ExStart:AddBidiMarks
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        builder->get_ParagraphFormat()->set_Bidi(true);
        builder->Writeln(u"שלום עולם!");
        builder->Writeln(u"مرحبا بالعالم!");

        auto saveOptions = MakeObject<TxtSaveOptions>();
        saveOptions->set_AddBidiMarks(true);

        doc->Save(ArtifactsDir + u"WorkingWithTxtSaveOptions.AddBidiMarks.txt", saveOptions);
        //ExEnd:AddBidiMarks
    }

    void UseTabCharacterPerLevelForListIndentation()
    {
        //ExStart:UseTabForListIndentation
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a list with three levels of indentation.
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->Writeln(u"Item 1");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->ListIndent();
        builder->Write(u"Item 3");

        auto saveOptions = MakeObject<TxtSaveOptions>();
        saveOptions->get_ListIndentation()->set_Count(1);
        saveOptions->get_ListIndentation()->set_Character(u'\t');

        doc->Save(ArtifactsDir + u"WorkingWithTxtSaveOptions.UseTabCharacterPerLevelForListIndentation.txt", saveOptions);
        //ExEnd:UseTabForListIndentation
    }

    void UseSpaceCharacterPerLevelForListIndentation()
    {
        //ExStart:UseSpaceForListIndentation
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a list with three levels of indentation.
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->Writeln(u"Item 1");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->ListIndent();
        builder->Write(u"Item 3");

        auto saveOptions = MakeObject<TxtSaveOptions>();
        saveOptions->get_ListIndentation()->set_Count(3);
        saveOptions->get_ListIndentation()->set_Character(u' ');

        doc->Save(ArtifactsDir + u"WorkingWithTxtSaveOptions.UseSpaceCharacterPerLevelForListIndentation.txt", saveOptions);
        //ExEnd:UseSpaceForListIndentation
    }

    void ExportHeadersFootersMode()
    {
        //ExStart:ExportHeadersFootersMode
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        auto doc = MakeObject<Document>();

        // Insert even and primary headers/footers into the document.
        // The primary header/footers will override the even headers/footers.
        doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, HeaderFooterType::HeaderEven));
        doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderEven)->AppendParagraph(u"Even header");
        doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, HeaderFooterType::FooterEven));
        doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterEven)->AppendParagraph(u"Even footer");
        doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, HeaderFooterType::HeaderPrimary));
        doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderPrimary)->AppendParagraph(u"Primary header");
        doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, HeaderFooterType::FooterPrimary));
        doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary)->AppendParagraph(u"Primary footer");

        // Insert pages to display these headers and footers.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Page 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Write(u"Page 3");

        auto options = MakeObject<TxtSaveOptions>();
        options->set_SaveFormat(SaveFormat::Text);

        // All headers and footers are placed at the very end of the output document.
        options->set_ExportHeadersFootersMode(TxtExportHeadersFootersMode::AllAtEnd);
        doc->Save(ArtifactsDir + u"WorkingWithTxtSaveOptions.HeadersFootersMode.AllAtEnd.txt", options);

        // Only primary headers and footers are exported at the beginning and end of each section.
        options->set_ExportHeadersFootersMode(TxtExportHeadersFootersMode::PrimaryOnly);
        doc->Save(ArtifactsDir + u"WorkingWithTxtSaveOptions.HeadersFootersMode.PrimaryOnly.txt", options);

        // No headers and footers are exported.
        options->set_ExportHeadersFootersMode(TxtExportHeadersFootersMode::None);
        doc->Save(ArtifactsDir + u"WorkingWithTxtSaveOptions.HeadersFootersMode.None.txt", options);
        //ExEnd:ExportHeadersFootersMode
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
