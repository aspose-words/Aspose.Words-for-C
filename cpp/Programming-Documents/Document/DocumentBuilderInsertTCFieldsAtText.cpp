#include "stdafx.h"
#include "examples.h"

#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <Model/Text/Range.h>
#include <Model/Text/Font.h>
#include <Model/Nodes/Node.h>
#include <Model/FindReplace/ReplacingArgs.h>
#include <Model/FindReplace/ReplaceAction.h>
#include <Model/FindReplace/IReplacingCallback.h>
#include <Model/FindReplace/FindReplaceOptions.h>
#include <Model/Fields/Field.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/DocumentBase.h>
#include <Model/Document/Document.h>
#include <drawing/color.h>
#include <cstdint>

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
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);
    public:
        InsertTCFieldHandler(const System::String& text, const System::String& switches)
            : mFieldText(text), mFieldSwitches(switches) {}

        InsertTCFieldHandler(const System::String& switches)
            : mFieldText(System::String::Empty), mFieldSwitches(switches) {}


        ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args) override
        {
            // Create a builder to insert the field.
            System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(System::DynamicCast<Aspose::Words::Document>(args->get_MatchNode()->get_Document()));
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
            return Aspose::Words::Replacing::ReplaceAction::Skip;
        }

    private:
        System::String mFieldText;
        System::String mFieldSwitches;
    };
	// ExEnd:InsertTCFieldHandler
}

void DocumentBuilderInsertTCFieldsAtText()
{
    // ExStart:DocumentBuilderInsertTCFieldsAtText
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    
    System::SharedPtr<FindReplaceOptions> options = System::MakeObject<FindReplaceOptions>();
    
    // Highlight newly inserted content.
    options->get_ApplyFont()->set_HighlightColor(System::Drawing::Color::get_DarkOrange());
    options->set_ReplacingCallback(System::MakeObject<InsertTCFieldHandler>(u"Chapter 1", u"\\l 1"));
    
    // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"The Beginning"), u"", options);
    // ExEnd:DocumentBuilderInsertTCFieldsAtText
}
