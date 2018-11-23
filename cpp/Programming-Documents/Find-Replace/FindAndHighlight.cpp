#include "stdafx.h"
#include "examples.h"

#include <system/text/regularexpressions/regex_options.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/enum_helpers.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <system/collections/ilist.h>

#include <Model/FindReplace/IReplacingCallback.h>
#include <Model/Text/Run.h>
#include <Model/Text/Range.h>
#include <Model/Text/Font.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/CompositeNode.h>
#include <Model/FindReplace/ReplacingArgs.h>
#include <Model/FindReplace/ReplaceAction.h>
#include <Model/FindReplace/IReplacingCallback.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/FindReplace/FindReplaceDirection.h>
#include <Model/Document/Document.h>
#include <drawing/color.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{

    // ExStart:SplitRun
    /// <summary>
    /// Splits text of the specified run into two runs.
    /// Inserts the new run just after the specified run.
    /// </summary>
    System::SharedPtr<Run> SplitRun(const System::SharedPtr<Run>& run, int32_t position)
    {
        auto afterRun = System::DynamicCast<Run>((System::StaticCast<Node>(run))->Clone(true));
        afterRun->set_Text(run->get_Text().Substring(position));
        run->set_Text(run->get_Text().Substring(0, position));
        run->get_ParentNode()->InsertAfter(afterRun, run);
        return afterRun;
    }
    // ExEnd:SplitRun

    // ExStart:ReplaceEvaluatorFindAndHighlight
    class ReplaceEvaluatorFindAndHighlight : public IReplacingCallback
    {
        RTTI_INFO(ReplaceEvaluatorFindAndHighlight, ::System::BaseTypesInfo<IReplacingCallback>);
    public:
        /// <summary>
        /// This method is called by the Aspose.Words find and replace engine for each match.
        /// This method highlights the match string, even if it spans multiple runs.
        /// </summary>
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> e) override;
    };

    ReplaceAction ReplaceEvaluatorFindAndHighlight::Replacing(System::SharedPtr<ReplacingArgs> e)
    {
        // This is a Run node that contains either the beginning or the complete match.
        System::SharedPtr<Node> currentNode = e->get_MatchNode();

        // The first (and may be the only) run can contain text before the match, 
        // In this case it is necessary to split the run.
        if (e->get_MatchOffset() > 0)
        {
            currentNode = SplitRun(System::DynamicCast<Run>(currentNode), e->get_MatchOffset());
        }

        // This array is used to store all nodes of the match for further highlighting.
        std::vector<System::SharedPtr<Run>> runs;

        // Find all runs that contain parts of the match string.
        int32_t remainingLength = e->get_Match()->get_Value().get_Length();
        while ((remainingLength > 0) && (currentNode != nullptr) && (currentNode->GetText().get_Length() <= remainingLength))
        {
            runs.push_back(System::DynamicCast<Run>(currentNode));
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
            auto run = System::DynamicCast<Run>(currentNode); 
            SplitRun(run, remainingLength);
            runs.push_back(run);
        }

        // Now highlight all runs in the sequence.
        for (auto& run : runs)
        {
            run->get_Font()->set_HighlightColor(System::Drawing::Color::get_Yellow());
        }

        // Signal to the replace engine to do nothing because we have already done all what we wanted.
        return ReplaceAction::Skip;
    }
    // ExEnd:ReplaceEvaluatorFindAndHighlight
}

void FindAndHighlight()
{
    std::cout << "FindAndHighlight example started." << std::endl;
    using namespace System::Text::RegularExpressions;
    // ExStart:FindAndHighlight
    // The path to the documents directory.
    System::String dataDir = GetDataDir_FindAndReplace();
    
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    
    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
    options->set_ReplacingCallback(System::MakeObject<ReplaceEvaluatorFindAndHighlight>());
    options->set_Direction(FindReplaceDirection::Backward);
    
    // We want the "your document" phrase to be highlighted.
    auto regex = System::MakeObject<Regex>(u"your document", RegexOptions::IgnoreCase);
    doc->get_Range()->Replace(regex, u"", options);
    
    System::String outputPath = dataDir + GetOutputFilePath(u"FindAndHighlight.doc");
    // Save the output document.
    doc->Save(outputPath);
    // ExEnd:FindAndHighlight
    std::cout << "Text highlighted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "FindAndHighlight example finished." << std::endl << std::endl;
}
