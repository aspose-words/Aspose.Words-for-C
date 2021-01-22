#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceDirection.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Story.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>

using namespace System::Text::RegularExpressions;
using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{
    void IgnoreTextInsideFields()
    {
        // ExStart:IgnoreTextInsideFields
        // Create document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert field with text inside.
        builder->InsertField(u"INCLUDETEXT", u"Text in field");

        System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();

        // Replace 'e' in document ignoring text inside field.
        options->set_IgnoreFields(true);
        doc->get_Range()->Replace(System::MakeObject<Regex>(u"e"), u"*", options);
        std::cout << doc->GetText().ToUtf8String() << std::endl; // The output is: \u0013INCLUDETEXT\u0014Text in field\u0015\f

        // Replace 'e' in document NOT ignoring text inside field.
        options->set_IgnoreFields(false);
        doc->get_Range()->Replace(System::MakeObject<Regex>(u"e"), u"*", options);
        std::cout << doc->GetText().ToUtf8String() << std::endl; // The output is: \u0013INCLUDETEXT\u0014T*xt in fi*ld\u0015\f
        // ExEnd:IgnoreTextInsideFields
    }

    void IgnoreTextInsideDeleteRevisions()
    {
        // ExStart:IgnoreTextInsideDeleteRevisions
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert non-revised text.
        builder->Writeln(u"Deleted");
        builder->Write(u"Text");

        // Remove first paragraph with tracking revisions.
        doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->Remove();
        doc->StopTrackRevisions();

        System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();

        // Replace 'e' in document while deleted text.
        options->set_IgnoreDeleted(true);
        doc->get_Range()->Replace(System::MakeObject<Regex>(u"e"), u"*", options);
        std::cout << doc->GetText().ToUtf8String() << std::endl; // The output is: Deleted\rT*xt\f

        // Replace 'e' in document NOT ignoring deleted text.
        options->set_IgnoreDeleted(false);
        doc->get_Range()->Replace(System::MakeObject<Regex>(u"e"), u"*", options);
        std::cout << doc->GetText().ToUtf8String() << std::endl; // The output is: D*l*t*d\rT*xt\f
        // ExEnd:IgnoreTextInsideDeleteRevisions
    }

    void IgnoreTextInsideInsertRevisions()
    {
        // ExStart:IgnoreTextInsideInsertRevisions
        // Create new document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert text with tracking revisions.
        doc->StartTrackRevisions(u"author", System::DateTime::get_Now());
        builder->Writeln(u"Inserted");
        doc->StopTrackRevisions();

        // Insert non-revised text.
        builder->Write(u"Text");

        System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();

        // Replace 'e' in document ignoring inserted text.
        options->set_IgnoreInserted(true);
        doc->get_Range()->Replace(System::MakeObject<Regex>(u"e"), u"*", options);
        std::cout << doc->GetText().ToUtf8String() << std::endl; // The output is: Inserted\rT*xt\f

        // Replace 'e' in document NOT ignoring inserted text.
        options->set_IgnoreInserted(false);
        doc->get_Range()->Replace(System::MakeObject<Regex>(u"e"), u"*", options);
        std::cout << doc->GetText().ToUtf8String() << std::endl; // The output is: Ins*rt*d\rT*xt\f
        // ExEnd:IgnoreTextInsideInsertRevisions
    }
}

void IgnoreText()
{
    std::cout << "IgnoreText examples started." << std::endl;
    using namespace System::Text::RegularExpressions;
    // ExStart:IgnoreText
    IgnoreTextInsideFields();
    IgnoreTextInsideDeleteRevisions();
    IgnoreTextInsideInsertRevisions();
    // ExEnd:IgnoreText
    std::cout << "IgnoreText examples finished." << std::endl << std::endl;
}
