#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/ImageData.h>
#include <Model/Drawing/Shape.h>
#include <Model/Markup/Sdt/SdtListItem.h>
#include <Model/Markup/Sdt/SdtListItemCollection.h>
#include <Model/Markup/Sdt/SdtType.h>
#include <Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/Node.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/Run.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Markup;

namespace
{
    void SetCurrentStateOfCheckBox(System::String const &dataDir)
    {
        // ExStart:SetCurrentStateOfCheckBox
        // Open an existing document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"CheckBoxTypeContentControl.docx");

        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        // Get the first content control from the document
        System::SharedPtr<StructuredDocumentTag> SdtCheckBox = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));

        // StructuredDocumentTag.Checked property gets/sets current state of the Checkbox SDT
        if (SdtCheckBox->get_SdtType() == SdtType::Checkbox)
        {
            SdtCheckBox->set_Checked(true);
        }

        System::String outputPath = dataDir + GetOutputFilePath(u"UpdateContentControls.SetCurrentStateOfCheckBox.docx");
        doc->Save(outputPath);
        // ExEnd:SetCurrentStateOfCheckBox
        std::cout << "Current state fo checkbox setup successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ModifyContentControls(System::String const &dataDir)
    {
        // ExStart:ModifyContentControls
        // Open an existing document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"CheckBoxTypeContentControl.docx");

        for (System::SharedPtr<StructuredDocumentTag> sdt : System::IterateOver<System::SharedPtr<StructuredDocumentTag>>(doc->GetChildNodes(NodeType::StructuredDocumentTag, true)))
        {
            if (sdt->get_SdtType() == SdtType::PlainText)
            {
                sdt->RemoveAllChildren();
                System::SharedPtr<Paragraph> para = System::DynamicCast_noexcept<Paragraph>(sdt->AppendChild(System::MakeObject<Paragraph>(doc)));
                System::SharedPtr<Run> run = System::MakeObject<Run>(doc, u"new text goes here");
                para->AppendChild(run);
            }
            else if (sdt->get_SdtType() == SdtType::DropDownList)
            {
                System::SharedPtr<SdtListItem> secondItem = sdt->get_ListItems()->idx_get(2);
                sdt->get_ListItems()->set_SelectedValue(secondItem);
            }
            else if (sdt->get_SdtType() == SdtType::Picture)
            {
                System::SharedPtr<Shape> shape = System::DynamicCast<Shape>(sdt->GetChild(NodeType::Shape, 0, true));
                if (shape->get_HasImage())
                {
                    shape->get_ImageData()->SetImage(dataDir + u"Watermark.png");
                }
            }
        }

        System::String outputPath = dataDir + GetOutputFilePath(u"UpdateContentControls.ModifyContentControls.docx");
        doc->Save(outputPath);
        // ExEnd:ModifyContentControls
        std::cout << "Plain text box, drop down list and picture content modified successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void UpdateContentControls()
{
    std::cout << "UpdateContentControls example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    SetCurrentStateOfCheckBox(dataDir);
    // Shows how to modify content controls of type plain text box, drop down list and picture.
    ModifyContentControls(dataDir);
    std::cout << "UpdateContentControls example finished." << std::endl << std::endl;
}