#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;
using namespace System::Text::RegularExpressions;

namespace
{
    // ExStart:ReplaceWithHtml
    class ReplaceWithHtmlEvaluator : public IReplacingCallback
    {
        typedef ReplaceWithHtmlEvaluator ThisType;
        typedef IReplacingCallback BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);

    public:
        ReplaceWithHtmlEvaluator(System::SharedPtr<FindReplaceOptions> options);
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> args) override;

    private:
        System::SharedPtr<FindReplaceOptions> mOptions;
    };

    ReplaceWithHtmlEvaluator::ReplaceWithHtmlEvaluator(System::SharedPtr<FindReplaceOptions> options)
    {
        mOptions = options;
    }

    ReplaceAction ReplaceWithHtmlEvaluator::Replacing(System::SharedPtr<ReplacingArgs> args)
    {
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(System::DynamicCast<Document>(args->get_MatchNode()->get_Document()));
        builder->MoveTo(args->get_MatchNode());

        // Replace '<CustomerName>' text with a red bold name.
        builder->InsertHtml(u"<b><font color='red'>James Bond, </font></b>");
        args->set_Replacement(u"");

        return ReplaceAction::Replace;
    }

    void ReplaceWithHtml(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello <CustomerName>,");

        auto options = System::MakeObject<FindReplaceOptions>();
        auto optionsReplacingCallback = System::MakeObject<ReplaceWithHtmlEvaluator>(options);

        doc->get_Range()->Replace(new Regex(u" <CustomerName>,"), System::String::Empty, options);

        // Save the modified document. 
        doc->Save(outputDataDir + u"Range.ReplaceWithInsertHtml.doc");
    }
    // ExEnd:ReplaceWithHtml

    // ExStart:NumberHighlightCallback
    // Replace and Highlight Numbers.
    class NumberHighlightCallback : public IReplacingCallback
    {
        typedef NumberHighlightCallback ThisType;
        typedef IReplacingCallback BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);
    public:
        NumberHighlightCallback(System::SharedPtr<FindReplaceOptions> const& opt)
            : mOpt(opt) { }

        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override
        {
            // Let replacement to be the same text.
            args->set_Replacement(args->get_Match()->get_Value());
            auto val = System::Convert::ToInt32(args->get_Match()->get_Value());

            // Apply either red or green color depending on the number value sign.
            mOpt->get_ApplyFont()->set_Color(val > 0 ? System::Drawing::Color::get_Green() : System::Drawing::Color::get_Red());
            return ReplaceAction::Replace;
        }
    private:
        System::SharedPtr<FindReplaceOptions> mOpt;
    };
    // ExEnd:NumberHighlightCallback

    // ExStart:LineCounter
    class LineCounterCallback : public IReplacingCallback
    {
        typedef LineCounterCallback ThisType;
        typedef IReplacingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);
    public:
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override
        {
            std::cout << args->get_Match()->get_Value().ToUtf8String() << '\n';
            args->set_Replacement(System::String::Format(u"{0} {1}", mCounter++, args->get_Match()->get_Value()));
            return ReplaceAction::Replace;
        }
    private:
        int32_t mCounter = 1;
    };

    void LineCounter(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // Create a document.
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);

        // Add lines of text.
        builder->Writeln(u"This is first line");
        builder->Writeln(u"Second line");
        builder->Writeln(u"And last line");

        // Prepend each line with line number.
        auto opt = System::MakeObject<FindReplaceOptions>();
        opt->set_ReplacingCallback(System::MakeObject<LineCounterCallback>());
        doc->get_Range()->Replace(System::MakeObject<Regex>(u"[^&p]*&p"), u"", opt);

        doc->Save(outputDataDir + u"TestLineCounter.docx");
    }
    // ExEnd:LineCounter

}

void ReplaceWithHTML()
{
    std::cout << "ReplaceWithHTML example started." << std::endl;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();

    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"Hello <CustomerName>,");

    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
    options->set_ReplacingCallback(System::MakeObject<ReplaceWithHtmlEvaluator>(options));

    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u" <CustomerName>,"), System::String::Empty, options);

    // Save the modified document.
    System::String outputPath = outputDataDir + u"ReplaceWithHTML.doc";
    doc->Save(outputPath);
    std::cout << "ReplaceWithHTML example finished." << std::endl << std::endl;
}