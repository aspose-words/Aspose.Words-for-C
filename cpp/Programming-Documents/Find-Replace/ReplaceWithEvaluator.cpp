#include "stdafx.h"
#include "examples.h"

#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/Range.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/FindReplace/ReplacingArgs.h>
#include <Model/FindReplace/ReplaceAction.h>
#include <Model/FindReplace/IReplacingCallback.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/Document/Document.h>
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
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> e) override
        {
            
            e->set_Replacement(e->get_Match()->ToString() + System::Convert::ToString(mMatchNumber));
            mMatchNumber++;
            return ReplaceAction::Replace;
        }
        
    private:
    
        int32_t mMatchNumber = 0;
    };
    // ExEnd:MyReplaceEvaluator
}

void ReplaceWithEvaluator()
{
    // ExStart:ReplaceWithEvaluator
    // The path to the documents directory.
    System::String dataDir = GetDataDir_FindAndReplace();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Range.ReplaceWithEvaluator.doc");
    
    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
    options->set_ReplacingCallback(System::MakeObject<::MyReplaceEvaluator>());
    
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"[s|m]ad"), u"", options);
    
    dataDir = dataDir + GetOutputFilePath(u"ReplaceWithEvaluator.doc");
    doc->Save(dataDir);
    // ExEnd:ReplaceWithEvaluator
    std::cout << "\nText replaced successfully with evaluator.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
