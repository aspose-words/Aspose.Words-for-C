#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{
    // ExStart:MyReplaceEvaluator
    class MyReplaceEvaluator : public IReplacingCallback
    {
        typedef MyReplaceEvaluator ThisType;
        typedef IReplacingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo)

    public:
        /// <summary>
        /// This is called during a replace operation each time a match is found.
        /// This method appends a number to the match string and returns it as a replacement string.
        /// </summary>
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> e) override;

    private:
        int32_t mMatchNumber = 0;
    };

    ReplaceAction MyReplaceEvaluator::Replacing(System::SharedPtr<ReplacingArgs> e)
    {
        e->set_Replacement(e->get_Match()->ToString() + System::Convert::ToString(mMatchNumber));
        mMatchNumber++;
        return ReplaceAction::Replace;
    }
    // ExEnd:MyReplaceEvaluator
}

void ReplaceWithEvaluator()
{
    std::cout << "ReplaceWithEvaluator example started." << std::endl;
    // ExStart:ReplaceWithEvaluator
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_FindAndReplace();
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Range.ReplaceWithEvaluator.doc");

    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
    options->set_ReplacingCallback(System::MakeObject<::MyReplaceEvaluator>());

    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"[s|m]ad"), u"", options);

    System::String outputPath = outputDataDir + u"ReplaceWithEvaluator.doc";
    doc->Save(outputPath);
    // ExEnd:ReplaceWithEvaluator
    std::cout << "Text replaced successfully with evaluator." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ReplaceWithEvaluator example finished." << std::endl << std::endl;
}
