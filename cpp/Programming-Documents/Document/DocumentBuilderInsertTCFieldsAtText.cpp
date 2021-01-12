#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <drawing/color.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace
{
    // ExStart:InsertTCFieldHandler
    class InsertTCFieldHandler : public IReplacingCallback
    {
        typedef InsertTCFieldHandler ThisType;
        typedef IReplacingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;

    public:
        InsertTCFieldHandler(const System::String& text, const System::String& switches)
            : mFieldText(text), mFieldSwitches(switches) {}
        InsertTCFieldHandler(const System::String& switches)
            : mFieldText(System::String::Empty), mFieldSwitches(switches) {}
        ReplaceAction Replacing(System::SharedPtr<ReplacingArgs> args) override;

    private:
        System::String mFieldText;
        System::String mFieldSwitches;
    };

    ReplaceAction InsertTCFieldHandler::Replacing(System::SharedPtr<ReplacingArgs> args)
    {
        // Create a builder to insert the field.
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(System::DynamicCast<Document>(args->get_MatchNode()->get_Document()));
        // Move to the first node of the match.
        builder->MoveTo(args->get_MatchNode());

        // If the user specified text to be used in the field as display text then use that, otherwise use the 
        // Match string as the display text.
        System::String insertText;

        if (!System::String::IsNullOrEmpty(mFieldText))
        {
            insertText = mFieldText;
        }
        else
        {
            insertText = args->get_Match()->get_Value();
        }

        // Insert the TC field before this node using the specified string as the display text and user defined switches.
        builder->InsertField(System::String::Format(u"TC \"{0}\" {1}", insertText, mFieldSwitches));

        // We have done what we want so skip replacement.
        return ReplaceAction::Skip;
    }
    // ExEnd:InsertTCFieldHandler
}

void DocumentBuilderInsertTCFieldsAtText()
{
    std::cout << "DocumentBuilderInsertTCFieldsAtText example started." << std::endl;
    // ExStart:DocumentBuilderInsertTCFieldsAtText
    System::SharedPtr<Document> doc = System::MakeObject<Document>();

    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();

    // Highlight newly inserted content.
    options->get_ApplyFont()->set_HighlightColor(System::Drawing::Color::get_DarkOrange());
    options->set_ReplacingCallback(System::MakeObject<InsertTCFieldHandler>(u"Chapter 1", u"\\l 1"));

    // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"The Beginning"), u"", options);
    // ExEnd:DocumentBuilderInsertTCFieldsAtText
    std::cout << "DocumentBuilderInsertTCFieldsAtText example finished." << std::endl << std::endl;
}
