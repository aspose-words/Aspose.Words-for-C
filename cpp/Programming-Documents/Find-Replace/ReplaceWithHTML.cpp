#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{
    class ReplaceWithHtmlEvaluator : public IReplacingCallback
    {
        typedef ReplaceWithHtmlEvaluator ThisType;
        typedef IReplacingCallback BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        ReplaceWithHtmlEvaluator(System::SharedPtr<FindReplaceOptions> options);
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> args);

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

    RTTI_INFO_IMPL_HASH(3215654403u, ReplaceWithHtmlEvaluator, ThisTypeBaseTypesInfo);
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