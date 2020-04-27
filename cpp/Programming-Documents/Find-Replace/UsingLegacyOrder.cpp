#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace System::Text::RegularExpressions;
using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{
    // ExStart:ReplaceUsingLegacyOrder
    class ReplacingCallback : public IReplacingCallback
    {
        typedef ReplacingCallback ThisType;
        typedef IReplacingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> e);
    };

    RTTI_INFO_IMPL_HASH(369195648u, ReplacingCallback, ThisTypeBaseTypesInfo);

    ReplaceAction ReplacingCallback::Replacing(System::SharedPtr<ReplacingArgs> e)
    {
        System::SharedPtr<Match> current = e->get_Match();

        //Signal to the replace engine to do nothing because we have already done all what we wanted.
        return ReplaceAction::Replace;
    }
    // ExEnd:ReplaceUsingLegacyOrder
}

void UsingLegacyOrder()
{
    std::cout << "UsingLegacyOrder example started." << std::endl;
    // ExStart:UsingLegacyOrder
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_FindAndReplace();
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
    options->set_ReplacingCallback(System::MakeObject<ReplacingCallback>());
    options->set_UseLegacyOrder(true);

    doc->get_Range()->Replace(System::MakeObject<Regex>(u"\\[(.* ? )\\]"), u"", options);

    System::String outputPath = outputDataDir + u"UsingLegacyOrder.doc";
    doc->Save(outputPath);
    // ExEnd:UsingLegacyOrder
    std::cout << "UsingLegacyOrder example finished." << std::endl << std::endl;
}