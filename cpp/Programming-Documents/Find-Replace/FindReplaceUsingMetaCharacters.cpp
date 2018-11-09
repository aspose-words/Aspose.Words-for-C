#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <Model/Text/Range.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/ParagraphAlignment.h>
#include <Model/Text/Font.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Document/BreakType.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{
    void MetaCharactersInSearchPattern(const System::String& dataDir)
    {
        // ExStart:MetaCharactersInSearchPattern
        // Initialize a Document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>();

        // Use a document builder to add content to the document.
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"This is Line 1");
        builder->Writeln(u"This is Line 2");

        auto findReplaceOptions = System::MakeObject<FindReplaceOptions>();

        doc->get_Range()->Replace(u"This is Line 1&pThis is Line 2", u"This is replaced line", findReplaceOptions);

        builder->MoveToDocumentEnd();
        builder->Write(u"This is Line 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"This is Line 2");

        doc->get_Range()->Replace(u"This is Line 1&mThis is Line 2", u"Page break is replaced with new text.", findReplaceOptions);

        auto savePath = dataDir + GetOutputFilePath(u"FindReplaceUsingMetaCharacters.MetaCharactersInSearchPattern.docx");
        doc->Save(savePath);
        // ExEnd:MetaCharactersInSearchPattern
        std::cout << "Find and Replace text using meta-characters has done successfully." << std::endl << "File saved at " << savePath.ToUtf8String() << std::endl;
}

    void ReplaceTextContaingMetaCharacters(const System::String& dataDir)
    {
        // ExStart:ReplaceTextContaingMetaCharacters
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Arial");
        builder->Writeln(u"First section");
        builder->Writeln(u"  1st paragraph");
        builder->Writeln(u"  2nd paragraph");
        builder->Writeln(u"{insert-section}");
        builder->Writeln(u"Second section");
        builder->Writeln(u"  1st paragraph");

        System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
        options->get_ApplyParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        // Double each paragraph break after word "section", add kind of underline and make it centered.
        int32_t count = doc->get_Range()->Replace(u"section&p", u"section&p----------------------&p", options);

        // Insert section break instead of custom text tag.
        count = doc->get_Range()->Replace(u"{insert-section}", u"&b", options);

        auto savePath = dataDir + GetOutputFilePath(u"FindReplaceUsingMetaCharacters.ReplaceTextContaingMetaCharacters.docx");
        doc->Save(savePath);
        // ExEnd:ReplaceTextContaingMetaCharacters
        std::cout << "Find and Replace text using meta-characters has done successfully." << std::endl << "File saved at " << savePath.ToUtf8String() << std::endl;
    }

}

void FindReplaceUsingMetaCharacters()
{
    /* meta-characters
                   &p - paragraph break
                   &b - section break
                   &m - page break
                   &l - manual line break */
    std::cout << "FindReplaceUsingMetaCharacters example started." << std::endl;
    // ExStart:FindReplaceUsingMetaCharacters
    // The path to the documents directory.
    System::String dataDir = GetDataDir_FindAndReplace();
    MetaCharactersInSearchPattern(dataDir);
    ReplaceTextContaingMetaCharacters(dataDir);
    // ExEnd:FindReplaceUsingMetaCharacters
    std::cout << "FindReplaceUsingMetaCharacters example finished." << std::endl << std::endl;
}
