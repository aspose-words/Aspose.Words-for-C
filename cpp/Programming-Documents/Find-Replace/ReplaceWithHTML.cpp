#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/FindReplace/IReplacingCallback.h>
#include <Model/FindReplace/ReplaceAction.h>
#include <Model/FindReplace/ReplacingArgs.h>
#include <Model/Text/Range.h>

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

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

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

    System::Object::shared_members_type ReplaceWithHtmlEvaluator::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("ReplaceWithHtmlEvaluator::mOptions", this->mOptions);
        return result;
    }


    RTTI_INFO_IMPL_HASH(3215654403u, ReplaceWithHtmlEvaluator, ThisTypeBaseTypesInfo);
}

void ReplaceWithHTML()
{
    std::cout << "ReplaceWithHTML example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_FindAndReplace();

    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"Hello <CustomerName>,");

    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
    options->set_ReplacingCallback(System::MakeObject<ReplaceWithHtmlEvaluator>(options));

    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u" <CustomerName>,"), System::String::Empty, options);

    // Save the modified document.
    System::String outputPath = dataDir + GetOutputFilePath(u"ReplaceWithHTML.doc");
    doc->Save(outputPath);
    std::cout << "ReplaceWithHTML example finished." << std::endl << std::endl;
}