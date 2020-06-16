#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Replacing;

namespace
{
    typedef System::SharedPtr<Node> TNodePtr;

    class ReplaceTextWithFieldHandler : public IReplacingCallback
    {
        typedef ReplaceTextWithFieldHandler ThisType;
        typedef IReplacingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);

    public:
        ReplaceTextWithFieldHandler(FieldType type);
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> args) override;
        std::vector<TNodePtr> FindAndSplitMatchRuns(System::SharedPtr<ReplacingArgs> args);
    private:
        FieldType mFieldType;

        System::SharedPtr<Run> SplitRun(System::SharedPtr<Run> run, int32_t position);
    };

    ReplaceTextWithFieldHandler::ReplaceTextWithFieldHandler(FieldType type) : mFieldType((FieldType)0)
    {
        mFieldType = type;
    }

    ReplaceAction ReplaceTextWithFieldHandler::Replacing(System::SharedPtr<ReplacingArgs> args)
    {
        std::vector<TNodePtr> runs = FindAndSplitMatchRuns(args);

        // Create DocumentBuilder which is used to insert the field.
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(System::DynamicCast<Document>(args->get_MatchNode()->get_Document()));
        builder->MoveTo(System::DynamicCast<Run>(runs.at(runs.size() - 1)));

        // Calculate the name of the field from the FieldType enumeration by removing the first instance of "Field" from the text. 
        // This works for almost all of the field types.
        System::String fieldName = System::ObjectExt::ToString(mFieldType).ToUpper().Substring(5);

        // Insert the field into the document using the specified field type and the match text as the field name.
        // If the fields you are inserting do not require this extra parameter then it can be removed from the string below.
        builder->InsertField(System::String::Format(u"{0} {1}",fieldName,args->get_Match()->get_Groups()->idx_get(0)));

        // Now remove all runs in the sequence.
        for (System::SharedPtr<Node> runNode : runs)
        {
            runNode->Remove();
        }

        // Signal to the replace engine to do nothing because we have already done all what we wanted.
        return ReplaceAction::Skip;
    }

    std::vector<TNodePtr> ReplaceTextWithFieldHandler::FindAndSplitMatchRuns(System::SharedPtr<ReplacingArgs> args)
    {
        // This is a Run node that contains either the beginning or the complete match.
        System::SharedPtr<Node> currentNode = args->get_MatchNode();

        // The first (and may be the only) run can contain text before the match, 
        // In this case it is necessary to split the run.
        if (args->get_MatchOffset() > 0)
        {
            currentNode = SplitRun(System::DynamicCast<Run>(currentNode), args->get_MatchOffset());
        }

        // This array is used to store all nodes of the match for further removing.
        std::vector<TNodePtr> runs;

        // Find all runs that contain parts of the match string.
        int32_t remainingLength = args->get_Match()->get_Value().get_Length();
        while ((remainingLength > 0) && (currentNode != nullptr) && (currentNode->GetText().get_Length() <= remainingLength))
        {
            runs.push_back(currentNode);
            remainingLength = remainingLength - currentNode->GetText().get_Length();

            // Select the next Run node. 
            // Have to loop because there could be other nodes such as BookmarkStart etc.
            do
            {
                currentNode = currentNode->get_NextSibling();
            } while ((currentNode != nullptr) && (currentNode->get_NodeType() != NodeType::Run));
        }

        // Split the last run that contains the match if there is any text left.
        if ((currentNode != nullptr) && (remainingLength > 0))
        {
            SplitRun(System::DynamicCast<Run>(currentNode), remainingLength);
            runs.push_back(currentNode);
        }

        return runs;
    }

    System::SharedPtr<Run> ReplaceTextWithFieldHandler::SplitRun(System::SharedPtr<Run> run, int32_t position)
    {
        System::SharedPtr<Run> afterRun = System::DynamicCast<Run>(System::StaticCast<Node>(run)->Clone(true));
        afterRun->set_Text(run->get_Text().Substring(position));
        run->set_Text(run->get_Text().Substring(0, position));
        run->get_ParentNode()->InsertAfter(afterRun, run);
        return afterRun;
    }
}

void ReplaceTextWithField()
{
    std::cout << "ReplaceTextWithField example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_FindAndReplace();
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Field.ReplaceTextWithFields.doc");

    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
    options->set_ReplacingCallback(System::MakeObject<ReplaceTextWithFieldHandler>(FieldType::FieldMergeField));

    // Replace any "PlaceHolderX" instances in the document (where X is a number) with a merge field.
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"PlaceHolder(\\d+)"), u"", options);

    System::String outputPath = outputDataDir + u"ReplaceTextWithField.doc";
    doc->Save(outputPath);

    std::cout << "Text replaced with field successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ReplaceTextWithField example finished." << std::endl << std::endl;
}