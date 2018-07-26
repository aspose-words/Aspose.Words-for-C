#include "../../examples.h"

#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/group_collection.h>
#include <system/text/regularexpressions/group.h>
#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/collections/ienumerator.h>
#include <Model/Text/Run.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/CompositeNode.h>
#include <Model/Fields/Nodes/FieldStart.h>
#include <Model/Fields/Nodes/FieldSeparator.h>
#include <Model/Fields/Nodes/FieldEnd.h>
#include <Model/Fields/FieldType.h>
#include <Model/Fields/Field.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace
{

    // ExStart:MergeField
    /// <summary>
    /// Represents a facade object for a merge field in a Microsoft Word document.
    /// </summary>
    class MergeField
    {
    public:
        /// <summary>
        /// Gets  the name of the merge field.
        /// </summary>
        System::String get_Name()
        {
            return (System::DynamicCast<Aspose::Words::Fields::FieldStart>(mFieldStart))->GetField()->get_Result().Replace(u"«", u"").Replace(u"»", u"");
        }

        /// <summary>
        /// Sets the name of the merge field.
        /// </summary>
        void set_Name(System::String value)
        {
            // Merge field name is stored in the field result which is a Run
            // Node between field separator and field end.
            auto fieldResult = System::DynamicCast<Aspose::Words::Run>(mFieldSeparator->get_NextSibling());
            fieldResult->set_Text(System::String::Format(u"«{0}»", value));

            // But sometimes the field result can consist of more than one run, delete these runs.
            RemoveSameParent(fieldResult->get_NextSibling(), mFieldEnd);

            UpdateFieldCode(value);
        }
        
        MergeField(const System::SharedPtr<FieldStart>& fieldStart)
            : mFieldStart(fieldStart)
        {
            if (!fieldStart)
            {
                throw System::ArgumentNullException(u"fieldStart");
            }
            if (fieldStart->get_FieldType() != Aspose::Words::Fields::FieldType::FieldMergeField)
            {
                throw System::ArgumentException(u"Field start type must be FieldMergeField.");
            }
            mFieldSeparator = fieldStart->GetField()->get_Separator();
            if (mFieldSeparator == nullptr)
            {
                throw System::InvalidOperationException(u"Cannot find field separator.");
            }
            mFieldEnd = fieldStart->GetField()->get_FieldEnd();

        }
    private:
        void UpdateFieldCode(System::String fieldName)
        {
            // Field code is stored in a Run node between field start and field separator.
            System::SharedPtr<Run> fieldCode = System::DynamicCast<Aspose::Words::Run>(mFieldStart->get_NextSibling());

            System::SharedPtr<System::Text::RegularExpressions::Match> match = gRegex().Match((System::DynamicCast<FieldStart>(mFieldStart))->GetField()->GetFieldCode());

            System::String newFieldCode = System::String::Format(u" {0}{1} ", match->get_Groups()->idx_get(u"start")->get_Value(), fieldName);
            fieldCode->set_Text(newFieldCode);

            // But sometimes the field code can consist of more than one run, delete these runs.
            RemoveSameParent(fieldCode->get_NextSibling(), mFieldSeparator);
        }

        /// <summary>
        /// Removes nodes from start up to but not including the end node.
        /// Start and end are assumed to have the same parent.
        /// </summary>
        static void RemoveSameParent(const System::SharedPtr<Node>& startNode, const System::SharedPtr<Node>& endNode)
        {
            if ((endNode != nullptr) && startNode->get_ParentNode() != endNode->get_ParentNode())
            {
                throw System::ArgumentException(u"Start and end nodes are expected to have the same parent.");
            }

            System::SharedPtr<Node> curChild = startNode;
            while ((curChild != nullptr) && (curChild != endNode))
            {
                System::SharedPtr<Node> nextChild = curChild->get_NextSibling();
                curChild->Remove();
                curChild = nextChild;
            }
        }

        static System::Text::RegularExpressions::Regex& gRegex()
        {
            static System::SharedPtr<System::Text::RegularExpressions::Regex> gRegex = System::MakeObject<System::Text::RegularExpressions::Regex>(u"\\s*(?<start>MERGEFIELD\\s|)(\\s|)(?<name>\\S+)\\s+");
            return *gRegex;
        }
    
        System::SharedPtr<Node> mFieldStart;
        System::SharedPtr<Node> mFieldSeparator;
        System::SharedPtr<Node> mFieldEnd;
    };
    // ExEnd:MergeField
}

void RenameMergeFields()
{
    // ExStart:RenameMergeFields
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();
    
    // Specify your document name here.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"RenameMergeFields.doc");
    
    // Select all field start nodes so we can find the merge fields.
    System::SharedPtr<NodeCollection> fieldStarts = doc->GetChildNodes(Aspose::Words::NodeType::FieldStart, true);
    
    auto fieldStart_enumerator = fieldStarts->GetEnumerator();
    System::SharedPtr<FieldStart> fieldStart;
    while (fieldStart_enumerator->MoveNext() && (fieldStart = System::DynamicCast<FieldStart>(fieldStart_enumerator->get_Current()), true))
    {
        if (fieldStart->get_FieldType() == Aspose::Words::Fields::FieldType::FieldMergeField)
        {
            MergeField mergeField{fieldStart};
            mergeField.set_Name(mergeField.get_Name() + u"_Renamed");
        }
    }
    
    dataDir = dataDir + GetOutputFilePath(u"RenameMergeFields.doc");
    doc->Save(dataDir);
    // ExEnd:RenameMergeFields
    std::cout << "\nMerge fields rename successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
