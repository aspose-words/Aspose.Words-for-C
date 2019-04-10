#include "stdafx.h"
#include "examples.h"

#include <Model/Document/DocumentBuilder.h>
#include <Model/Fields/Fields/DocumentAutomation/FieldIf.h>
#include <Model/Fields/Fields/DocumentAutomation/FieldIfComparisonResult.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void EvaluateIFCondition()
{
    std::cout << "EvaluateIFCondition example started." << std::endl;
    // ExStart:EvaluateIFCondition
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>();
    System::SharedPtr<FieldIf> field = System::DynamicCast<FieldIf>(builder->InsertField(u"IF 1 = 1", nullptr));

    FieldIfComparisonResult actualResult = field->EvaluateCondition();
    std::cout << System::ObjectExt::Box<FieldIfComparisonResult>(actualResult)->ToString().ToUtf8String() << std::endl;
    // ExEnd:EvaluateIFCondition

    std::cout << "Evaluates the IF condition successfully." << std::endl;
    std::cout << "EvaluateIFCondition example finished." << std::endl << std::endl;
}