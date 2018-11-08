#pragma once

#include <system/shared_ptr.h>

namespace Aspose { namespace Words { class CompositeNode; } }
namespace Aspose { namespace Words { namespace Fields { enum class FieldType; } } }

void ConvertFieldsToStaticText(System::SharedPtr<Aspose::Words::CompositeNode> compositeNode, Aspose::Words::Fields::FieldType targetFieldType);