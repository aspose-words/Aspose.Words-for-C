#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Model/Text/Run.h>
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
        MergeField(const System::SharedPtr<FieldStart>& fieldStart);
        /// <summary>
        /// Gets  the name of the merge field.
        /// </summary>
        System::String GetName() { return (System::DynamicCast<FieldStart>(mFieldStart))->GetField()->get_Result().Replace(u"«", u"").Replace(u"»", u""); }
        /// <summary>
        /// Sets the name of the merge field.
        /// </summary>
        void SetName(const System::String &value);

    private:
        void UpdateFieldCode(const System::String &fieldName);
        /// <summary>
        /// Removes nodes from start up to but not including the end node.
        /// Start and end are assumed to have the same parent.
        /// </summary>
        static void RemoveSameParent(const System::SharedPtr<Node>& startNode, const System::SharedPtr<Node>& endNode);
        static System::Text::RegularExpressions::Regex& gRegex();
    
        System::SharedPtr<Node> mFieldStart;
        System::SharedPtr<Node> mFieldSeparator;
        System::SharedPtr<Node> mFieldEnd;
    };

    MergeField::MergeField(const System::SharedPtr<FieldStart>& fieldStart) : mFieldStart(fieldStart)
    {
        if (!fieldStart)
        {
            throw System::ArgumentNullException(u"fieldStart");
        }
        if (fieldStart->get_FieldType() != FieldType::FieldMergeField)
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

    void MergeField::SetName(const System::String &value)
    {
        // Merge field name is stored in the field result which is a Run
        // Node between field separator and field end.
        System::SharedPtr<Run> fieldResult = System::DynamicCast<Run>(mFieldSeparator->get_NextSibling());
        fieldResult->set_Text(System::String::Format(u"«{0}»", value));

        // But sometimes the field result can consist of more than one run, delete these runs.
        RemoveSameParent(fieldResult->get_NextSibling(), mFieldEnd);

        UpdateFieldCode(value);
    }

    void MergeField::UpdateFieldCode(const System::String &fieldName)
    {
        // Field code is stored in a Run node between field start and field separator.
        System::SharedPtr<Run> fieldCode = System::DynamicCast<Run>(mFieldStart->get_NextSibling());

        System::SharedPtr<System::Text::RegularExpressions::Match> match = gRegex().Match((System::DynamicCast<FieldStart>(mFieldStart))->GetField()->GetFieldCode());

        System::String newFieldCode = System::String::Format(u" {0}{1} ", match->get_Groups()->idx_get(u"start")->get_Value(), fieldName);
        fieldCode->set_Text(newFieldCode);

        // But sometimes the field code can consist of more than one run, delete these runs.
        RemoveSameParent(fieldCode->get_NextSibling(), mFieldSeparator);
    }

    void MergeField::RemoveSameParent(const System::SharedPtr<Node>& startNode, const System::SharedPtr<Node>& endNode)
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

    System::Text::RegularExpressions::Regex& MergeField::gRegex()
    {
        static System::SharedPtr<System::Text::RegularExpressions::Regex> gRegex = System::MakeObject<System::Text::RegularExpressions::Regex>(u"\\s*(?<start>MERGEFIELD\\s|)(\\s|)(?<name>\\S+)\\s+");
        return *gRegex;
    }
    // ExEnd:MergeField
}

void RenameMergeFields()
{
    std::cout << "RenameMergeFields example started." << std::endl;
    // ExStart:RenameMergeFields
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();

    // Specify your document name here.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"RenameMergeFields.doc");

    // Select all field start nodes so we can find the merge fields.
    System::SharedPtr<NodeCollection> fieldStarts = doc->GetChildNodes(NodeType::FieldStart, true);

    for (System::SharedPtr<FieldStart> fieldStart : System::IterateOver<System::SharedPtr<FieldStart>>(fieldStarts))
    {
        if (fieldStart->get_FieldType() == FieldType::FieldMergeField)
        {
            MergeField mergeField{fieldStart};
            mergeField.SetName(mergeField.GetName() + u"_Renamed");
        }
    }

    System::String outputPath = outputDataDir + u"RenameMergeFields.doc";
    doc->Save(outputPath);
    // ExEnd:RenameMergeFields
    std::cout << "Merge fields rename successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RenameMergeFields example finished." << std::endl << std::endl;
}
