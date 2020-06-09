#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUpdateCultureProvider.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldUpdateCultureSource.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace
{
    typedef System::SharedPtr<System::Globalization::CultureInfo> TCultureInfoPtr;

    class FieldUpdateCultureProvider : public IFieldUpdateCultureProvider
    {
        typedef FieldUpdateCultureProvider ThisType;
        typedef IFieldUpdateCultureProvider BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        TCultureInfoPtr GetCulture(System::String name, System::SharedPtr<Field> field) override;
    };

    RTTI_INFO_IMPL_HASH(2889170660u, FieldUpdateCultureProvider, ThisTypeBaseTypesInfo);

    TCultureInfoPtr FieldUpdateCultureProvider::GetCulture(System::String name, System::SharedPtr<Field> field)
    {
        if (name == u"ru-RU")
        {
            TCultureInfoPtr culture = System::MakeObject<System::Globalization::CultureInfo>(name, false);
            System::SharedPtr<System::Globalization::DateTimeFormatInfo> format = culture->get_DateTimeFormat();
            format->set_MonthNames(System::MakeArray<System::String>({u"мес€ц 1", u"мес€ц 2", u"мес€ц 3", u"мес€ц 4", u"мес€ц 5", u"мес€ц 6", u"мес€ц 7", u"мес€ц 8", u"мес€ц 9", u"мес€ц 10", u"мес€ц 11", u"мес€ц 12", u""}));
            format->set_MonthGenitiveNames(format->get_MonthNames());
            format->set_AbbreviatedMonthNames(System::MakeArray<System::String>({u"мес 1", u"мес 2", u"мес 3", u"мес 4", u"мес 5", u"мес 6", u"мес 7", u"мес 8", u"мес 9", u"мес 10", u"мес 11", u"мес 12", u""}));
            format->set_AbbreviatedMonthGenitiveNames(format->get_AbbreviatedMonthNames());
            format->set_DayNames(System::MakeArray<System::String>({u"день недели 7", u"день недели 1", u"день недели 2", u"день недели 3", u"день недели 4", u"день недели 5", u"день недели 6"}));
            format->set_AbbreviatedDayNames(System::MakeArray<System::String>({u"день 7", u"день 1", u"день 2", u"день 3", u"день 4", u"день 5", u"день 6"}));
            format->set_ShortestDayNames(System::MakeArray<System::String>({u"д7", u"д1", u"д2", u"д3", u"д4", u"д5", u"д6"}));
            format->set_AMDesignator(u"ƒо полудн€");
            format->set_PMDesignator(u"ѕосле полудн€");
            const System::String pattern = u"yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
            format->set_LongDatePattern(pattern);
            format->set_LongTimePattern(pattern);
            format->set_ShortDatePattern(pattern);
            format->set_ShortTimePattern(pattern);
            return culture;
        }
        else if (name == u"en-US")
        {
            return System::MakeObject<System::Globalization::CultureInfo>(name, false);
        }
        else
        {
            return nullptr;
        }
    }
}

void FieldUpdateCulture()
{
    std::cout << "FieldUpdateCulture example started." << std::endl;
    // ExStart:FieldUpdateCultureProvider
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"FieldUpdateCultureProvider.docx");

    doc->get_FieldOptions()->set_FieldUpdateCultureSource(FieldUpdateCultureSource::FieldCode);
    doc->get_FieldOptions()->set_FieldUpdateCultureProvider(System::MakeObject<FieldUpdateCultureProvider>());

    System::String outputPath = outputDataDir + u"FieldUpdateCulture.pdf";
    doc->Save(outputPath);
    // ExEnd:FieldUpdateCultureProvider

    std::cout << "Format of Time field is set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "FieldUpdateCulture example finished." << std::endl << std::endl;
}