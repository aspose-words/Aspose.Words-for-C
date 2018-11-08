#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <system/collections/ienumerator.h>
#include <system/collections/idictionary.h>
#include <system/collections/dictionary.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/ListFormat.h>
#include <Model/Sections/SectionStart.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Lists/ListCollection.h>
#include <Model/Lists/List.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Lists;

void ListUseDestinationStyles()
{
    std::cout << "ListUseDestinationStyles example started." << std::endl;
    // ExStart:ListUseDestinationStyles
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.DestinationList.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.SourceList.doc");

    // Set the source document to continue straight after the end of the destination document.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(Aspose::Words::SectionStart::Continuous);

    // Keep track of the lists that are created.
    System::SharedPtr<System::Collections::Generic::IDictionary<int32_t, System::SharedPtr<List>>> newLists = System::MakeObject<System::Collections::Generic::Dictionary<int32_t, System::SharedPtr<List>>>();

    // Iterate through all paragraphs in the document.
    auto para_enumerator = srcDoc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->GetEnumerator();
    System::SharedPtr<Paragraph> para;
    while (para_enumerator->MoveNext() && (para = System::DynamicCast<Paragraph>(para_enumerator->get_Current()), true))
    {
        if (para->get_IsListItem())
        {
            int32_t listId = para->get_ListFormat()->get_List()->get_ListId();

            // Check if the destination document contains a list with this ID already. If it does then this may
            // Cause the two lists to run together. Create a copy of the list in the source document instead.
            if (dstDoc->get_Lists()->GetListByListId(listId) != nullptr)
            {
                System::SharedPtr<List> currentList;
                // A newly copied list already exists for this ID, retrieve the stored list and use it on 
                // The current paragraph.
                if (newLists->ContainsKey(listId))
                {
                    currentList = newLists->idx_get(listId);
                }
                else
                {
                    // Add a copy of this list to the document and store it for later reference.
                    currentList = srcDoc->get_Lists()->AddCopy(para->get_ListFormat()->get_List());
                    newLists->Add(listId, currentList);
                }

                // Set the list of this paragraph  to the copied list.
                para->get_ListFormat()->set_List(currentList);
            }
        }
    }

    // Append the source document to end of the destination document.
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::UseDestinationStyles);

    System::String outputPath = dataDir + GetOutputFilePath(u"ListUseDestinationStyles.doc");
    // Save the combined document to disk.
    dstDoc->Save(outputPath);
    // ExEnd:ListUseDestinationStyles
    std::cout << "Document appended successfully without continuing any list numberings." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ListUseDestinationStyles example finished." << std::endl << std::endl;
}
