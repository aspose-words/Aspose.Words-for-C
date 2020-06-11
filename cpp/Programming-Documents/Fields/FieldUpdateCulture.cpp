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
            format->set_MonthNames(System::MakeArray<System::String>({u"����� 1", u"����� 2", u"����� 3", u"����� 4", u"����� 5", u"����� 6", u"����� 7", u"����� 8", u"����� 9", u"����� 10", u"����� 11", u"����� 12", u""}));
            format->set_MonthGenitiveNames(format->get_MonthNames());
            format->set_AbbreviatedMonthNames(System::MakeArray<System::String>({u"��� 1", u"��� 2", u"��� 3", u"��� 4", u"��� 5", u"��� 6", u"��� 7", u"��� 8", u"��� 9", u"��� 10", u"��� 11", u"��� 12", u""}));
            format->set_AbbreviatedMonthGenitiveNames(format->get_AbbreviatedMonthNames());
            format->set_DayNames(System::MakeArray<System::String>({u"���� ������ 7", u"���� ������ 1", u"���� ������ 2", u"���� ������ 3", u"���� ������ 4", u"���� ������ 5", u"���� ������ 6"}));
            format->set_AbbreviatedDayNames(System::MakeArray<System::String>({u"���� 7", u"���� 1", u"���� 2", u"���� 3", u"���� 4", u"���� 5", u"���� 6"}));
            format->set_ShortestDayNames(System::MakeArray<System::String>({u"�7", u"�1", u"�2", u"�3", u"�4", u"�5", u"�6"}));
            format->set_AMDesignator(u"�� �������");
            format->set_PMDesignator(u"����� �������");
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