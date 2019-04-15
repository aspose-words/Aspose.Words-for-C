#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/FindReplace/IReplacingCallback.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/FindReplace/ReplaceAction.h>
#include <Model/FindReplace/ReplacingArgs.h>
#include <Model/Nodes/Node.h>
#include <Model/Text/Range.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{
    // ExStart:ReplaceHtmlFindAndInsertHtml
    class FindAndInsertHtml : public IReplacingCallback
    {
        typedef FindAndInsertHtml ThisType;
        typedef IReplacingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> e);
    };

    RTTI_INFO_IMPL_HASH(369195648u, FindAndInsertHtml, ThisTypeBaseTypesInfo);

    ReplaceAction FindAndInsertHtml::Replacing(System::SharedPtr<ReplacingArgs> e)
    {
        // This is a Run node that contains either the beginning or the complete match.
        System::SharedPtr<Node> currentNode = e->get_MatchNode();

        // create Document Buidler and insert MergeField
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(System::DynamicCast_noexcept<Document>(e->get_MatchNode()->get_Document()));

        builder->MoveTo(currentNode);
        builder->InsertHtml(e->get_Replacement());
        currentNode->Remove();

        //Signal to the replace engine to do nothing because we have already done all what we wanted.
        return ReplaceAction::Skip;
    }
    // ExEnd:ReplaceHtmlFindAndInsertHtml
}

void ReplaceHtmlTextWithMetaCharacters()
{
    std::cout << "ReplaceHtmlTextWithMetaCharacters example started." << std::endl;
    // ExStart:ReplaceHtmlTextWithMetaCharacters
    // The path to the documents directory.
    System::String dataDir = GetDataDir_FindAndReplace();

    System::String html = u"<p>&ldquo;Some Text&rdquo;</p>";

    // Initialize a Document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();

    // Use a document builder to add content to the document.
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->Write(u"{PLACEHOLDER}");

    System::SharedPtr<FindReplaceOptions> findReplaceOptions = System::MakeObject<FindReplaceOptions>();
    findReplaceOptions->set_ReplacingCallback(System::MakeObject<FindAndInsertHtml>());
    findReplaceOptions->set_PreserveMetaCharacters(true);

    doc->get_Range()->Replace(u"{PLACEHOLDER}", html, findReplaceOptions);


    System::String outputPath = dataDir + GetOutputFilePath(u"ReplaceHtmlTextWithMetaCharacters.doc");
    doc->Save(outputPath);
    // ExEnd:ReplaceHtmlTextWithMetaCharacters
    std::cout << "Text replaced with meta characters successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ReplaceHtmlTextWithMetaCharacters example finished." << std::endl << std::endl;
}