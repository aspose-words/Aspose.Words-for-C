#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <system/text/string_builder.h>
#include <system/text/regularexpressions/regex.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;
using namespace System::Text::RegularExpressions;

namespace
{
    void ReplaceTextInFooter(System::String const& inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ReplaceTextInFooter
        // Open the template document, containing obsolete copyright information in the footer.
        auto doc = System::MakeObject<Document>(inputDataDir + u"HeaderFooter.ReplaceText.doc");

        // Access header of the Word document.
        auto headersFooters = doc->get_FirstSection()->get_HeadersFooters();
        auto header = headersFooters->idx_get(HeaderFooterType::HeaderPrimary);

        // Set options.
        auto options = System::MakeObject<FindReplaceOptions>();
        options->set_MatchCase(false);
        options->set_FindWholeWordsOnly(false);

        // Replace text in the header of the Word document.
        header->get_Range()->Replace(u"Aspose.Words", u"Remove", options);

        auto footer = headersFooters->idx_get(HeaderFooterType::FooterPrimary);
        footer->get_Range()->Replace(u"(C) 2006 Aspose Pty Ltd.", u"Copyright (C) Aspose Pty Ltd.", options);


        // Save the Word document.
        doc->Save(outputDataDir + u"HeaderReplace.docx");
        // ExEnd:ReplaceTextInFooter
    }

    // ExStart:ShowChangesForHeaderAndFooterOrders
    class ReplaceLog : public IReplacingCallback
    {
        using ThisType = ReplaceLog;
        using BaseType = IReplacingCallback;
        using ThisTypeBaseTypesInfo = ::System::BaseTypesInfo<BaseType>;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);
    public:
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override
        {
            mTextBuilder.AppendLine(args->get_MatchNode()->GetText());
            return ReplaceAction::Skip;
        }

        void ClearText()
        {
            mTextBuilder.Clear();
        }

        System::String Text() const
        {
            return mTextBuilder.ToString();
        }
    private:
        System::Text::StringBuilder mTextBuilder;
    };

    void ShowChangesForHeaderAndFooterOrders(System::String const &outputDataDir, System::String const& inputDataDir)
    {
        auto doc = System::MakeObject<Document>(inputDataDir + u"HeaderFooter.HeaderFooterOrder.docx");

        // Assert that we use special header and footer for the first page
        // The order for this: first header\footer, even header\footer, primary header\footer
        auto firstPageSection = doc->get_FirstSection();

        auto logger = System::MakeObject<ReplaceLog>();
        auto options = System::MakeObject<FindReplaceOptions>();
        options->set_ReplacingCallback(logger);

        doc->get_Range()->Replace(System::MakeObject<Regex>(u"(header|footer)") , u"", options);

        doc->Save(outputDataDir + u"HeaderFooter.HeaderFooterOrder.docx");

        // Prepare our string builder for assert results without "DifferentFirstPageHeaderFooter"
        logger->ClearText();

        // Remove special first page
        // The order for this: primary header, default header, primary footer, default footer, even header\footer
        firstPageSection->get_PageSetup()->set_DifferentFirstPageHeaderFooter(false);

        doc->get_Range()->Replace(System::MakeObject<Regex>(u"(header|footer)") , u"", options);

    }
    // ExEnd:ShowChangesForHeaderAndFooterOrders
}

void ReplaceInHeaderAndFooter()
{
    
    std::cout << "ReplaceInHeaderAndFooter example started." << std::endl;

    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_FindAndReplace();
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();

    ReplaceTextInFooter(inputDataDir, outputDataDir);
    ShowChangesForHeaderAndFooterOrders(inputDataDir, outputDataDir);
}

