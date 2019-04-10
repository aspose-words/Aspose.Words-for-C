#include "stdafx.h"
#include "examples.h"

#include <Model/Text/ControlChar.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Document/ConvertUtil.h>

using namespace Aspose::Words;

namespace
{
    void ConvertBetweenMeasurementUnits()
    {
        // ExStart:ConvertBetweenMeasurementUnits
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<PageSetup> pageSetup = builder->get_PageSetup();
        pageSetup->set_TopMargin(ConvertUtil::InchToPoint(1.0));
        pageSetup->set_BottomMargin(ConvertUtil::InchToPoint(1.0));
        pageSetup->set_LeftMargin(ConvertUtil::InchToPoint(1.5));
        pageSetup->set_RightMargin(ConvertUtil::InchToPoint(1.5));
        pageSetup->set_HeaderDistance(ConvertUtil::InchToPoint(0.2));
        pageSetup->set_FooterDistance(ConvertUtil::InchToPoint(0.2));
        // ExEnd:ConvertBetweenMeasurementUnits
        std::cout << "Page properties specified in inches." << std::endl;
}

    void UseControlCharacters()
    {
        // ExStart:UseControlCharacters
        System::String text = u"test\r";
        // Replace "\r" control character with "\r\n"
        text = text.Replace(ControlChar::Cr(), ControlChar::CrLf());
        // ExEnd:UseControlCharacters
        std::cout << "nControl characters used successfully." << std::endl;
    }
}

void UtilityClasses()
{
    std::cout << "UtilityClasses example started." << std::endl;
    ConvertBetweenMeasurementUnits();
    UseControlCharacters();
    std::cout << "UtilityClasses example finished." << std::endl << std::endl;
}
